#!/usr/bin/python
import os
import sys
import string
import pyodbc
import re

# Dosage Guidelines
def PredictDosageLevel(statin, dosage):
	if statin == "atorva":
		if dosage >= 10:
			if dosage >= 40:
				return "high"
			return "moder."
		return "low"
	elif statin == "rosuva":
		if dosage >= 5:
			if dosage >= 20:
				return "high"
			return "moder."
		return "low"
	elif statin == "simva":
		if dosage >= 20:
			if dosage >= 80:
				return "high"
			return "moder."
		return "low"
	elif statin == "prava":
		if dosage >= 40:
			return "moder."
		return "low"
	elif statin == "lova":
		if dosage >= 40:
			return "moder."
		return "low"
	elif statin == "fluva":
		if dosage >= 80:
			return "moder."
		return "low"
	elif statin == "pitava":
		if dosage >= 2:
			return "moder."
		return "low"
	#errors
	return ""

#readIn currentId and update ID file
NodeIDNum = 0
with open('id', 'r') as f:
	NodeIDNum = f.readline()
	NodeIDNum = int(NodeIDNum)
	NodeIDNum += 1
	
#Connection to Access DB
conn1 = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=***PATH TO DB***\ChartReviewDB.accdb')
c1 = conn1.cursor()
conn2 = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=***PATH TO ADDTIONAL PATIENT DATA***\QUIPS Intervention.accdb')
c2 = conn2.cursor()

### PROCESSING KEY WORDS###
categoryBools = {}
keyWords = {}
statinWords = {}
statinExact = {}
keyWordsList = []
scores = {}

#initialize scores
scores[1] = 0
scores[2] = 0
scores[3] = 0
scores[4] = 0
scores[5] = 0
scores[6] = 0
scores[7] = 0
scores[8] = 0
scores[9] = 0
scores[10] = 0
scores[11] = 0

#statins keywords 
keyWordFile = open('Key_Words\statin_terms.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	statinWords[temp] = 0
keyWordFile.close()

#statins types
keyWordFile = open('Key_Words\statins_exact.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	statinExact[temp] = 0
keyWordFile.close()

#cat1 'Outside provider'
keyWordFile = open('Key_Words\keyWordsCat1.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	keyWordsList.append(temp)
	keyWords[temp] = 1
keyWordFile.close()

#cat2 'On statins'
keyWordFile = open('Key_Words\keyWordsCat2.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	keyWordsList.append(temp)
	keyWords[temp] = 2
keyWordFile.close()

#cat3 'Patient refusal'
keyWordFile = open('Key_Words\keyWordsCat3.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	keyWordsList.append(temp)
	keyWords[temp] = 3
keyWordFile.close()

#cat4 'Started statin'
keyWordFile = open('Key_Words\keyWordsCat4.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	keyWordsList.append(temp)
	keyWords[temp] = 4
keyWordFile.close()

#cat5 'No time during visit'
keyWordFile = open('Key_Words\keyWordsCat5.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	keyWordsList.append(temp)
	keyWords[temp] = 5
keyWordFile.close()

#cat6 'Life Expectancy < 5 years'
keyWordFile = open('Key_Words\keyWordsCat6.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	keyWordsList.append(temp)
	keyWords[temp] = 6
keyWordFile.close()

#cat7 'Allergies'
keyWordFile = open('Key_Words\keyWordsCat7.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	keyWordsList.append(temp)
	keyWords[temp] = 7
keyWordFile.close()

#cat8 'I do not think he needs one'
keyWordFile = open('Key_Words\keyWordsCat8.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	keyWordsList.append(temp)
	keyWords[temp] = 8
keyWordFile.close()

#cat9 'Patient wants to stay on statin'
keyWordFile = open('Key_Words\keyWordsCat9.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	keyWordsList.append(temp)
	keyWords[temp] = 9
keyWordFile.close()

#cat10 'Patient is not taking a statin'
keyWordFile = open('Key_Words\keyWordsCat10.txt', 'r')
for word in keyWordFile:
	temp = word.rstrip()
	keyWordsList.append(temp)
	keyWords[temp] = 10
keyWordFile.close()

### END PROCESSING KEY WORDS ###

### NOTE PROCESSING ###
count = 1
for file in os.listdir('Notes'):
	print ('Processing Note: ' + str(count))
	globalStatinBool = 0
	poppedWords = []
	poppedSections = []
	statinWordsFound = []
	statinSections = []
	addedStatinTerms = {}
	fullNote = ""
	section = ""
	dosage = 0
	distToDos = 25
	dosBool = 0
	statinType = ""
	statinAlone = 0
	statinDosageLevel = ""
	statusError =  "No"
	
	# Process note data for searching
	# Rebuild original note while key words identified
	with open('Notes\\' + file) as note:
		ssn = (note.name[6:])
		ssn = ssn[:9]
		### START Loop through each line in note
		for line in note:
			fullNote += line + '<Br>'
			
			if line not in ['\n', '\r\n']:
				section += line + '<Br>'
			else:
				section = section[:-1]
				sectionOrig = section
				section = section.lower()
				statinBool = 0
				poppedString = ""
				
				################################
				##### START KEY WORD LOOP ######
				################################
				
				### START Loop through each word in keywords
				for word in keyWords:
					if word in section:
						if word == "start":
							if "not" in section:
								if (section.find(word) - section.find("not")) < 5:
									continue
						### START Loop through each statin in statinWords
						for statin in statinWords:
							if statin in section:
								if statin == "statin":
									continue
								if statin not in addedStatinTerms.keys():
									poppedString += ' <b> ' + statin + ' </b> '
									addedStatinTerms[statin] = 1
								sectionOrig = sectionOrig.lower().replace(statin, '<b>' + statin + '</b>')
								fullNote = fullNote.replace(statin, '<b>' + statin + '</b>')
								globalStatinBool = 1
								statinBool = 1
								# Search for instances of direct mention of key words
								## Score based on distance to closest mention of statin term
								startLocs = [match.start() for match in re.finditer(re.escape(word), section)]
								for x in startLocs:
									distance = abs(x - section.find(statin))
									if distance >= 40:
										scores[keyWords[word]] += 5
									else:
										scores[keyWords[word]] += ((40 - distance) + 5)
								# If statin is called out non-generically
								if statin != "statin":
									# Search around exact statin term for dosage level
									y = 1
									### START Loop - Determine dosage
									while y < 100:
										distance = section.find(str(y)) - section.find(statin)
										if ((distance < 25) and (distance <= distToDos) and (distance > 0)):
											dosBool = 1
											dosage = y
											distToDos = distance
											# Half-Tablet Check
											if (section.find("one-half") != -1):
												if abs(section.find("one-half") - section.find(str(y))) < 20:
														dosage /= 2
										if y == 1:
											y = 2
										elif y == 2:
											y = 5
										else:
											y += 5
									### END Loop - Determine dosage
						### END Loop through each word in keywords
						
						poppedString += ' <b> ' + word + ' </b> ' 
						sectionOrig = sectionOrig.lower().replace(word, '<b>' + word + '</b>')
						fullNote = fullNote.replace(word, '<b>' + word + '</b>')
				### END Loop through each word in keywords
				
				###########################
				##### END KEY WORD LOOP ###
				###########################
				
				# Key term AND statin term found
				if statinBool == 1:
					statinWordsFound.append(poppedString)
					statinSections.append(sectionOrig)
				
				################################
				##### START STATIN TERM LOOP ###
				################################
				
				# Only statin term(s) found
				else:
					for statin in statinWords:
						if statin in section:
							# Ensure mention found isn't actually 'stating'
							# Alternative to searching "statin " <- extra space
							if statin == "statin":
								if section.find(statin) == section.find("stating"):
									continue
							statinBool = 1
							globalStatinBool = 1
							if statin != "statin":
								# Predicts dosage
								y = 1
								while y < 100:
									distance = section.find(str(y)) - section.find(statin)
									if ((distance < 25) and distance <= (distToDos) and (distance > 0)):
										dosBool = 1
										dosage = y
										distToDos = distance
										# Half-Tablet Check
										if (section.find("one-half") != -1):
											if abs((section.find("one-half") - section.find(str(y)))) < 20:
												dosage = y / 2
									if y == 1:
										y = 2
									elif y == 2:
										y = 5
									else:
										y += 5
							#assumes on statin and simply listed as a med
							##manual score bump
							scores[2] += 20
							poppedString += '<b>' + statin + '</b>'
							statinWordsFound.append(poppedString)
							sectionOrig = sectionOrig.lower().replace(statin, '<b> ' + statin + ' </b>')
							statinSections.append(sectionOrig)
							if statin != "statin":
								break
				
				##############################
				##### END STATIN TERM LOOP ###
				##############################
				
				section = ""
		### END Loop through each line in note
		
		# Checks if only one type of statin is in note
		#	if so, predicts dosage level
		for x in statinExact:
			if x in fullNote.lower():
				statinAlone += 1
		if statinAlone == 1 and dosBool == 1:
			for x in statinExact:
				if x in fullNote.lower():
					statinType = x
					statinDosageLevel = PredictDosageLevel(x, dosage)
					break
		else:
			statinAlone = 0
	### END NOTE PROCESSING ###

	### DETERMINING CATEGORY ###
	top = 0
	cat = 12
	if scores[1] > 0 and scores[1] > top:
		top = scores[1]
		cat = 1
	if scores[2] > 0 and scores[2] > top:
		top = scores[2]
		if statinDosageLevel == "low":
			cat = 2
		elif statinDosageLevel == "moder.":
			cat = 3
		elif statinDosageLevel == "high":
			cat = 3
		else:
			cat = 5
	if scores[3] > 0 and scores[3] > top:
		top = scores[3]
		cat = 6
	if scores[4] > 0 and scores[4] > top:
		top = scores[4]
		cat = 9
	if scores[5] > 0 and scores[5] > top:
		top = scores[5]
		cat = 10
	if scores[6] > 0 and scores[6] > top:
		top = scores[6]
		cat = 11
	if scores[7] > 0 and scores[7] > top:
		top = scores[7]
		cat = 12
	if scores[8] > 0 and scores[8] > top:
		top = scores[8]
		cat = 7
	if scores[9] > 0 and scores[9] > top:
		top = scores[9]
		cat = 8
	if scores[10] > 0 and scores[10] > top:
		top = scores[10]
		cat = 7
	### END DETERMINING CATEGORY

	### ACCESS DB STATEMENTS (EITHER OR BOTH)
	#Formats dosage
	dosStr = str(dosage)
	dosStr = dosStr + "mg"
	#Converts category scores to yes/no values and determine potential conflicts
	categoryBools = {}
	categoryBools['NoStatin'] = 0
	categoryBools['OutsideProvider'] = 0
	categoryBools['OnStatinsLow'] = 0
	categoryBools['OnStatinsMH'] = 0
	categoryBools['OnStatins'] = 0
	categoryBools['PatientRefusal'] = 0
	categoryBools['StartedStatins'] = 0
	categoryBools['NoTime'] = 0
	categoryBools['LifeExp'] = 0
	categoryBools['NoRec'] = 0
	categoryBools['PatStay'] = 0
	categoryBools['NoStatin'] = 0
	categoryBools['Allergy'] = 0
	
	if scores[1] > 10:
		categoryBools['OutsideProvider'] = 1
	if scores[2] > 10:
		if statinDosageLevel == "low":
			categoryBools['OnStatinsLow'] = 1
		elif statinDosageLevel == "moder.":
			categoryBools['OnStatinsMH'] = 1
		elif statinDosageLevel == "high":
			categoryBools['OnStatinsMH'] = 1
		else:
			categoryBools['OnStatins'] = 1
	if scores[3] > 10:
		categoryBools['PatientRefusal'] = 1
	if scores[4] > 10:
		categoryBools['StartedStatins'] = 1
	if scores[5] > 10:
		categoryBools['NoTime'] = 1
	if scores[6] > 10:
		categoryBools['LifeExp'] = 1
	if scores[7] > 10:
		categoryBools['Allergy'] = 1
	if scores[8] > 10:
		categoryBools['NoRec'] = 1
	if scores[9] > 10:
		categoryBools['PatStay'] = 1
	if scores[10] > 10:
		categoryBools['NoStatin'] = 1
	
	#Insert Note Data
	executeString = "INSERT INTO Notes VALUES (?, ?, ?, ?,\
	0, 0, '',\
	?, ?, ?, ?, ?, ?,\
	?, ?, ?, ?,\
	?, ?, ?, ?, ?, ?,\
	'',\
	?,\
	?, ?, ?,\
	?, ?,\
	0,\
	?, ?, ?,\
	?, 0)"
	""" Trouble Shooting 
	print (str(scores[1])+ str(scores[2])+ str(scores[3])+ str(scores[4])+ str(scores[5]), str(scores[6]))
	print (str(scores[8])+ str(scores[9])+ str(scores[10])+ str(scores[11]))
	print (str(globalStatinBool)+ dosStr+ str(dosBool)+ str(statinAlone)+ statinType+ statinDosageLevel)
	print (str(categoryBools['OutsideProvider']))
	print (str(categoryBools['OnStatinsLow'])+ str(categoryBools['OnStatinsMH'])+ str(categoryBools['OnStatins']))
	print (str(categoryBools['PatientRefusal'])+ str(categoryBools['NoRec']))
	print (str(categoryBools['StartedStatins'])+ str(categoryBools['NoTime'])+ str(categoryBools['LifeExp']))
	"""
	
	c1.execute(executeString, str(NodeIDNum), str(ssn), str(cat), str(fullNote),\
	str(scores[1]), str(scores[2]), str(scores[3]), str(scores[4]), str(scores[5]), str(scores[6]),\
	str(scores[8]), str(scores[9]), str(scores[10]), str(scores[11]),\
	str(globalStatinBool), dosStr, str(dosBool), str(statinAlone), statinType, statinDosageLevel,\
	str(categoryBools['OutsideProvider']),\
	str(categoryBools['OnStatinsLow']), str(categoryBools['OnStatinsMH']), str(categoryBools['OnStatins']),\
	str(categoryBools['PatientRefusal']), str(categoryBools['NoRec']),\
	str(categoryBools['StartedStatins']), str(categoryBools['NoTime']), str(categoryBools['LifeExp']), \
	str(categoryBools['Allergy']))

	#Insert Sections Data
	##With Statins
	numFound = len(statinWordsFound)
	x = 0
	while x < numFound:
		c1.execute("INSERT INTO  SectionsStatins VALUES (?, ?, ?)", str(NodeIDNum), statinSections[x], statinWordsFound[x])
		x += 1
	
	#Query for additional patient data
	c2.execute("SELECT TOP 1 PatientFirstName, patientlastname, score, statinlevel FROM all_pts WHERE patientssn = (?)", str(ssn))
	result = c2.fetchall()
	if len(result):
		temp = []
		for x in result[0]:
			temp.append(x)
		c1.execute("INSERT INTO Pt_Data VALUES (?, ?, ?, ?, ?)", str(ssn), str(temp[0]), str(temp[1]), str(temp[2]), str(temp[3]))
	else:
		errString = "Error: SSN not found in patient DB: " + str(ssn)
		print (errString)
		
	# Save results
	conn1.commit()
	conn2.commit()
	
	# increment ID
	NodeIDNum += 1
	count +=1
	### END ACCESS DB STATEMENTS
	
	#reset scores
	scores[1] = 0
	scores[2] = 0
	scores[3] = 0
	scores[4] = 0
	scores[5] = 0
	scores[6] = 0
	scores[7] = 0
	scores[8] = 0
	scores[9] = 0
	scores[10] = 0
	scores[11] = 0
	
	categoryBools['OutsideProvider'] = 0
	categoryBools['OnStatins'] = 0
	categoryBools['OnStatinsLow'] = 0
	categoryBools['OnStatinsMH'] = 0
	categoryBools['PatientRefusal'] = 0
	categoryBools['StartedStatins'] = 0
	categoryBools['NoTime'] = 0
	categoryBools['LifeExp'] = 0
	categoryBools['HealthStatus'] = 0
	categoryBools['NoRec'] = 0
	categoryBools['PatStay'] = 0
	categoryBools['NoStatin'] = 0
	categoryBools['Allergy'] = 0
	
#Update ID File
with open('id', 'r+') as f:
	f.seek(0)
	f.write(str(NodeIDNum))
print ('Complete.') 


