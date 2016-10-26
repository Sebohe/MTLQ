"""
Author: Sebastian Bolanos Heston, 2016, TopTier, LLC
About: Simple program to find all of the part numbers in a BOM that begin
with "PT11" or "EL17" and compare them to a lookup table with the motor specifications.
"""
from dbfread import DBF
import easygui
import os
import subprocess
#from pwd import getpwuid
from operator import itemgetter
#import glob
import sys
from shutil import copyfile
#import threading
#import tim

#Returns all fo the part numbers that belong in a BOM.
#This is a recursive function so it will look inside of assemblies
def returnALLPartNumbers(bomNumber,TABLE_DBF):
	templist1 = []
	for record in TABLE_DBF:
		if record['BOMNO'] == bomNumber:
			if record['PART_ASSY'] == 'P':
				templist1.append(record)
			elif record['PART_ASSY'] == 'A':
				#When it is an aassembly the part number is the next
				#bom number
				for item in returnALLPartNumbers(record['PARTNO'], TABLE_DBF):
					templist1.append(item)

	return templist1

def writePartNumbers(lista):
	f = open('partsList_Output','w+')
	for x in lista:
		f.write(x['PARTNO'])
		f.write(', ')
		try:
			f.write(x['REFDESMEMO'])
		except:
			f.write('NONE')
		f.write('\n')
	f.close()

def generateMotorList(partsList, PART_TABLE, bomNumber):
	manufacturer, model, motorMechanicaName, realPartNumberList = [],[],[],[]
	description, secondPartNumberList, templist1, largeMotorList,notFoundInMotorDictList = [],[],[],[],[]

	#Parses the parts list of the BOM and finds all of the PT11 and EL17 parts.
	for record in partsList:
		if 'PT11' in record['PARTNO'] or 'EL17' in record['PARTNO']:
			realPartNumberList.append(record['PARTNO'])
			if record['REFDESMEMO']:
				motorMechanicaName.append(record['REFDESMEMO'])
			else:
				motorMechanicaName.append('Name is an empty String. Tell mechanical to clean up!')


	lengthOfQuery = (len(realPartNumberList))
	#print (realPartNumberList)

	#Look in the MRP PART database. Find model number and addtional information.
	for part in realPartNumberList:
		for record in PART_TABLE:
			if record['PARTNO'] == part:
				secondPartNumberList.append(record['PARTNO'])
				manufacturer.append(record['MANUFACTER'])
				model.append(record['MODELNO'])
				description.append(record['DESCRIPT'])
				break


	#since the two tables saved at different indexes the part number and the mechanical name
	#needs to match the table list.
	#This compares the secondPartNumberList and the realPartNumberList
	#and looks for the same value.
	#This is no longer needed since the parts list is in the same order.
	"""
	for x in range(lengthOfQuery):
		for y in range(lengthOfQuery):
			if secondPartNumberList[x] == realPartNumberList[y]:
				if motorMechanicaName[y]:
					templist1.append(motorMechanicaName[y])
				else:
					templist1.append('MOTOR NAME IS EMPTY IN BOM')
	"""

	#opens and reads the contents of the motor lookup table.
	try:
		f = open('motors','r')
		lines = f.readlines()
		f.close()
		lineLenght = len(lines)
		lineCounter = 1
		for line in lines:
			x = line[0]
			x = line[1]
			x = line[2]
			x = line[3]
			lineCounter += 1
	except Exception as e:
		easygui.msgbox(str(e)+'\nLine: '+str(lineCounter))
		sys.exit()
	#comapares the description and model to the lookup table
	#to see if they match. When it does create a dictionary
	#with all the motor values then add that dictionary to a list.

	for x in range(lengthOfQuery):
		lineCounter = 0
		for line in lines:
			lineParameters = line.split(",")
			lineCounter += 1
			#print (description[x].strip())


			if  lineParameters[0].replace(" ","") in model[x] and lineParameters[1].replace(" ","") in description[x].replace(" ",""):
				individualMotorDict = {'Name' : motorMechanicaName[x],
										'Part':secondPartNumberList[x],
										'Manufacturer':manufacturer[x],
										'Model':model[x],
										'Description':description[x],
										'HP':lineParameters[1].strip(),
										'RPM':lineParameters[2].strip(),
										'A':lineParameters[3].rstrip('\n')}
				largeMotorList.append(individualMotorDict)

				break
			elif(lineLenght-lineCounter == 0):
				#'Name' : 'zName Not found for Part Number',

				individualMotorDict = {'Name' : motorMechanicaName[x],
										'Part':secondPartNumberList[x],
										'Manufacturer':manufacturer[x],
										'Model':model[x],
										'Description':description[x],
										'HP':'',
										'RPM':'',
										'A':''}
				notFoundInMotorDictList.append(individualMotorDict)


	#print (lineCounter)



	#Sort the motor by the motor name.
	newlist = sorted(largeMotorList, key=itemgetter('Name'))
	largeMotorList = newlist


	newlist = sorted(notFoundInMotorDictList, key=itemgetter('Name'))
	notFoundInMotorDictList = newlist

	#thread.join()
	return largeMotorList, notFoundInMotorDictList

def writeToOutput(foundMotorList, notFoundInMotorList,bomName):
	#opens a new file called output and deletes it's contents.
	f = open(bomName+'.txt','w+')
	counter = 0
	if foundMotorList:
		for motor in foundMotorList:
			counter = 1 + counter
			f.write("\n")
			f.write("\n")
			f.write(str(counter))
			f.write("\n")
			f.write("Name: ")
			f.write(motor['Name'].rstrip("\n"))
			f.write("\n")


			f.write("Part Number: ")
			f.write(motor['Part'])
			f.write("\n")

			f.write("Manufacturer: ")
			f.write(motor['Manufacturer'])
			f.write("\n")

			f.write("Model: ")
			f.write(motor['Model'])
			f.write("\n")

			f.write("Description: ")
			f.write(motor['Description'])
			f.write("\n")

			f.write("HP: ")
			f.write(motor['HP'])
			f.write("\n")

			f.write("RPM: ")
			f.write(motor['RPM'])
			f.write("\n")

			f.write("Amps: ")
			f.write(motor['A'])
			f.write("\n")
			f.write("______________________")
	else:
		f.write("The bom number was not found.")


	if notFoundInMotorList:
		f.write('\n')
		f.write('*********************\n\n')
		f.write('Motors not found in motor file. Some of them might need to be added to the motors file:\n\n')


		for motor in notFoundInMotorList:

			counter = 1 + counter
			f.write("\n")
			f.write("\n")
			f.write(str(counter))
			f.write("\n")
			f.write("Name: ")
			f.write(motor['Name'].rstrip("\n"))
			f.write("\n")


			f.write("Part Number: ")
			f.write(motor['Part'])
			f.write("\n")

			f.write("Manufacturer: ")
			f.write(motor['Manufacturer'])
			f.write("\n")

			f.write("Model: ")
			f.write(motor['Model'])
			f.write("\n")

			f.write("Description: ")
			f.write(motor['Description'])
			f.write("\n")

			f.write("HP: ")
			f.write(motor['HP'])
			f.write("\n")

			f.write("RPM: ")
			f.write(motor['RPM'])
			f.write("\n")

			f.write("Amps: ")
			f.write(motor['A'])
			f.write("\n")
			f.write("______________________")



	f.close()

	#calls notepad.exe to view the output.
	program = 'notepad.exe'
	fileName = bomName+'.txt'
	subprocess.Popen([program, fileName])

def getBomName(bomNumber, BOM_TABLE):
	for record in BOM_TABLE:
		if record['BOMNO'] == bomNumber:
		  	return (record['BOMDESCRI'])


def find_owner(filename):
    return getpwuid(stat(filename).st_uid).pw_name

if __name__=="__main__":
	easygui.msgbox('After OK, the program will load database to RAM. This can take a minute.\n\nIf you want the most up to date database, delete the database files located in the same path as this program.')
	MRPBOMpath = ['T:\\pcmrpw ver. 8.0\\MRPBOM.DBF','T:\\pcmrpw ver. 8.0\\mrpbom.fpt']
	MRPPARTpath = ['T:\\pcmrpw ver. 8.0\\MRPPART.DBF','T:\\pcmrpw ver. 8.0\\MRPPART.FPT']

	currentDir = os.path.dirname(os.path.realpath(__file__))
	#if 'MRPBOM.DBF' not in os.listdir(currentDir) and 'mrpbom.fpt' not in os.listdir(currentDir):
	#	copyfile(MRPBOMpath[0],currentDir+'\\MRPBOM.DBF')
	#	copyfile(MRPBOMpath[1],currentDir+'\\mrpbom.fpt')
	#if 'MRPPART.DBF' not in os.listdir(currentDir) and 'MRPPART.FPT' not in os.listdir(currentDir):
	#	copyfile(MRPPARTpath[0],currentDir+'\\MRPPART.DBF')
	#	copyfile(MRPPARTpath[1],currentDir+'\\MRPPART.FPT')
	#print('Loading database to RAM. This can take a minute.')
	try:
		BOM_TABLE = DBF(MRPBOMpath[0],load=True)
	except Exception as e:
		#print (os.stat())
		easygui.msgbox(str(e))
		sys.exit()


	try:
		PART_TABLE = DBF(MRPPARTpath[0],load=True)
	except Exception as e:
		#print (os.stat(MRPPARTpath[0]))
		easygui.msgbox(str(e))
		sys.exit()


	#print (find_owner(MRPPARTpath[0]))

	sys.setrecursionlimit(2000)


	#print('PART TABLE HEADERS')
	#print (PART_TABLE.field_names)
	#print()
	#print (PART_TABLE.field_names)
	#print('BOM TABLE HEADERS')
	#print (BOM_TABLE.field_names)
	while True:
		bomNumber = easygui.enterbox(msg="Enter the BOM you want to query.")

		#print (bomNumber[0])
		#print (type(bomNumber))
		if bomNumber[0] == "e":
			bomNumber = bomNumber.replace("e","E")

		if not bomNumber:
			sys.exit()

		partsList = returnALLPartNumbers(bomNumber, BOM_TABLE)

		if partsList:
			break
		else:
			if not easygui.ynbox(msg="No BOM found with the number: "+bomNumber+"\n\nTry again?"):
				sys.exit()

	#writePartNumbers(partsList)
	motorsList = generateMotorList(partsList,PART_TABLE,bomNumber)
	writeToOutput(motorsList[0],motorsList[1],getBomName(bomNumber,BOM_TABLE))
