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

def returnALLPartNumber(bomNumber,TABLE_DBF):
	templist1 = []
	for record in TABLE_DBF:
		if record['BOMNO'] == bomNumber:
			if record['PART_ASSY'] == 'P':
				templist1.append(record['PARTNO'])
			elif record['PART_ASSY'] == 'A':
				#When it is an aassem
				templist2 = returnALLPartNumber(record['PARTNO'], TABLE_DBF)
				for item in templist2:
					templist1.append(item)
	return templist1


def generateMotorList():

	MRPBOMpath = ['T:\\pcmrpw ver. 8.0\\MRPBOM.DBF','T:\\pcmrpw ver. 8.0\\mrpbom.fpt']
	MRPPARTpath = ['T:\\pcmrpw ver. 8.0\\MRPPART.DBF','T:\\pcmrpw ver. 8.0\\MRPPART.FPT']
	manufacturer = []
	model = []
	motorMechanicaName = []
	realPartNumberList = []
	description = []
	secondPartNumberList = []
	templist1 = []
	largeMotorList = []
	currentDir = os.path.dirname(os.path.realpath(__file__))
	bomNumber = easygui.enterbox(msg="Enter the BOM you want to look for.")
	if bomNumber == None:
		#easygui.msgbox("Exiting program.")
		sys.exit()

	# try to implemented the ifile capabilities that the dbfread library has here. Just to learn.

	if 'MRPBOM.DBF' not in os.listdir(currentDir) and 'mrpbom.fpt' not in os.listdir(currentDir):
		copyfile(MRPBOMpath[0],currentDir+'\\MRPBOM.DBF')
		copyfile(MRPBOMpath[1],currentDir+'\\mrpbom.fpt')
	if 'MRPPART.DBF' not in os.listdir(currentDir) and 'MRPPART.FPT' not in os.listdir(currentDir):
		copyfile(MRPPARTpath[0],currentDir+'\\MRPPART.DBF')
		copyfile(MRPPARTpath[1],currentDir+'\\MRPPART.FPT')

	#thread = threading.Thread(target=(easygui.msgbox("Please wait... Parsing through the database.")))
	#thread.start()
	#Loads the database
	print (time.clock())
	try:
		BOM_TABLE = DBF('MRPBOM.DBF', load=True)
		PART_TABLE = DBF('MRPPART.DBF', load=True)
	except Exception as e:
		easygui.msgbox(e)
	print (time.clock())

	#Parses through the BOM database in search PT in the partnumber.
	#If found get the motor description that mechanical gives it.
	#and get the actual PARTNUMBER
	for record in BOM_TABLE:
		if record['BOMNO'] == bomNumber:
			if 'PT' in record['PARTNO']:
				realPartNumberList.append(record['PARTNO'])
				motorMechanicaName.append(record['REFDESMEMO'])

	lengthOfQuery = (len(realPartNumberList))

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
				templist1.append(motorMechanicaName[y])

	motorMechanicaName = templist1


	#opens and reads the contents of the motor lookup table.
	f = open('motors','r')
	lines = f.readlines()
	f.close()
	#comapares the description and model to the lookup table
	#to see if they match. When it does create a dictionary
	#with all the motor values then add that dictionary to a list.
	for line in lines:
		lineList = line.split(",")
		for x in range(lengthOfQuery):
			if  lineList[0] in model[x] and lineList[1] in description[x]:
				individualMotorDict = {'Name' : motorMechanicaName[x],
										'Part':secondPartNumberList[x],
										'Manufacturer':manufacturer[x],
										'Model':model[x],
										'Description':description[x],
										'HP':lineList[1],
										'RPM':lineList[2],
										'A':lineList[3].rstrip('\n')}
				largeMotorList.append(individualMotorDict)

	#Sort the motor by the motor name.
	newlist = sorted(largeMotorList, key=itemgetter('Name'))
	largeMotorList = newlist
	print (time.clock())
	#thread.join()
	print (time.clock())
	return largeMotorList

def writeToOutput(motorDictList):
	#opens a new file called output and deletes it's contents.
	f = open('output','w+')
	counter = 0
	for motor in motorDictList:
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

	writeToOutput(generateMotorList())
