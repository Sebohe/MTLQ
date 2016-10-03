"""
Author: Sebastian Bolaños Heston, 2016, TopTier, LLC
About: Simple program to find all of the part numbers in a BOM that begin
with "PT" and compare them to a lookup table with the motor specifications.
"""
from dbfread import DBF
import easygui
import os
import subprocess
from operator import itemgetter
import glob
import sys
from shutil import copyfile
import threading
import time

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

	#Parses through the BOM database in search PT in the partnumber.
	#If found get the motor description that mechanical gives it.
	#and get the actual PARTNUMBER
	#bomNumber = easygui.enterbox(msg="Enter the BOM you want to look for.")
	for record in partsList:
		#if record['BOMNO'] == bomNumber:
		if 'PT11' in record['PARTNO'] or 'EL17' in record['PARTNO']:
			realPartNumberList.append(record['PARTNO'])
			motorMechanicaName.append(record['REFDESMEMO'])

	lengthOfQuery = (len(realPartNumberList))
	#print (realPartNumberList)
	#print (PART_TABLE.table_names)
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
	for x in range(lengthOfQuery):
		for y in range(lengthOfQuery):
			if secondPartNumberList[x] == realPartNumberList[y]:
				if motorMechanicaName[y]:
					templist1.append(motorMechanicaName[y])
				else:
					templist1.append('None')

	motorMechanicaName = templist1


	#opens and reads the contents of the motor lookup table.
	f = open('motors','r')
	lines = f.readlines()
	f.close()
	lineLenght = len(lines)
	#comapares the description and model to the lookup table
	#to see if they match. When it does create a dictionary
	#with all the motor values then add that dictionary to a list.

	for x in range(lengthOfQuery):
		lineCounter = 0
		for line in lines:
			lineParameters = line.split(",")
			lineCounter += 1


			if  lineParameters[0] in model[x] and lineParameters[1] in description[x].strip():
				individualMotorDict = {'Name' : motorMechanicaName[x],
										'Part':secondPartNumberList[x],
										'Manufacturer':manufacturer[x],
										'Model':model[x],
										'Description':description[x],
										'HP':lineParameters[1],
										'RPM':lineParameters[2],
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

def writeToOutput(foundMotorList, notFoundInMotorList):
	#opens a new file called output and deletes it's contents.
	f = open('output','w+')
	counter = 0
	for motor in foundMotorList:
		counter = 1 + counter
		f.write(str(counter))
		f.write("\n")
		f.write("Name: ")
		f.write(motor['Name'])
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
		f.write("\n")

	f.write('\n')
	f.write('*********************\n\n')
	f.write('Motors not found in motor file. Some of them might need to be added to the motors file:\n\n')
	for motor in notFoundInMotorList:
		counter = 1 + counter
		f.write(str(counter))
		f.write("\n")
		f.write("Name: ")
		f.write(motor['Name'])
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
		f.write("\n")

	f.close()

	#calls notepad.exe to view the output.
	program = 'notepad.exe'
	fileName = 'output'
	subprocess.Popen([program, fileName])

if __name__=="__main__":
	easygui.msgbox('RUNNING!')
	MRPBOMpath = ['T:\\pcmrpw ver. 8.0\\MRPBOM.DBF','T:\\pcmrpw ver. 8.0\\mrpbom.fpt']
	MRPPARTpath = ['T:\\pcmrpw ver. 8.0\\MRPPART.DBF','T:\\pcmrpw ver. 8.0\\MRPPART.FPT']

	currentDir = os.path.dirname(os.path.realpath(__file__))
	if 'MRPBOM.DBF' not in os.listdir(currentDir) and 'mrpbom.fpt' not in os.listdir(currentDir):
		copyfile(MRPBOMpath[0],currentDir+'\\MRPBOM.DBF')
		copyfile(MRPBOMpath[1],currentDir+'\\mrpbom.fpt')
	if 'MRPPART.DBF' not in os.listdir(currentDir) and 'MRPPART.FPT' not in os.listdir(currentDir):
		copyfile(MRPPARTpath[0],currentDir+'\\MRPPART.DBF')
		copyfile(MRPPARTpath[1],currentDir+'\\MRPPART.FPT')
	#print('Loading database to RAM. This can take a minute.')
	try:
		BOM_TABLE = DBF('MRPBOM.DBF',load=True)
		PART_TABLE = DBF('MRPPART.DBF',load=True)
	except Exception as e:
		easygui.msgbox(e)

	sys.setrecursionlimit(2000)
	bomNumber = easygui.enterbox(msg="Enter the BOM you want to query.")
	if not bomNumber:
		sys.exit()

	#print('PART TABLE HEADERS')
	#print (PART_TABLE.field_names)
	#print('BOM TABLE HEADERS')
	#print (BOM_TABLE.field_names)
	partsList = returnALLPartNumbers(bomNumber, BOM_TABLE)
	writePartNumbers(partsList)
	motorsList = generateMotorList(partsList,PART_TABLE,bomNumber)
	writeToOutput(motorsList[0],motorsList[1])
