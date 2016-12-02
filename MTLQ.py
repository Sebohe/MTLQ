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
def returnALLRecords(bomNumber,TABLE_DBF):
	templist1 = []
	for record in TABLE_DBF:
		if record['BOMNO'] == bomNumber:
			if record['PART_ASSY'] == 'P':
				templist1.append(record)
			elif record['PART_ASSY'] == 'A':
				#When it is an aassembly the part number is the next
				#bom number
				for item in returnALLRecords(record['PARTNO'], TABLE_DBF):
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

def generateMotorList(partsList, PART_TABLE):
	manufacturer, model, motorMechanicaName, realPartNumberList = [],[],[],[]
	quantity, assemblyNumber, assembly, description, secondPartNumberList, templist1, largeMotorList,notFoundInMotorDictList = [],[],[],[],[],[],[],[]


	#Parses the parts list of the BOM and finds all of the PT11 and EL17 parts.
	for record in partsList:
		if 'PT11' in record['PARTNO'] or 'EL17' in record['PARTNO']:
			realPartNumberList.append(record['PARTNO'])
			if record['REFDESMEMO']:
				motorMechanicaName.append(record['REFDESMEMO'])
				assembly.append(record['BOMDESCRI'])
				assemblyNumber.append(record['BOMNO'])
				quantity.append(record['QTY'])
			else:
				motorMechanicaName.append('Name is an empty String. Tell mechanical to clean up!')
				assembly.append(record['BOMDESCRI'])
				assemblyNumber.append(record['BOMNO'])
				quantity.append(record['QTY'])


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
										'Assembly':assembly[x],
										'Part':secondPartNumberList[x],
										'Manufacturer':manufacturer[x],
										'AssemblyNumber': assemblyNumber[x],
										'Model':model[x],
										'Description':description[x],
										'QTY': quantity[x],
										'HP':lineParameters[1].strip(),
										'RPM':lineParameters[2].strip(),
										'A':lineParameters[3].rstrip('\n')}
				largeMotorList.append(individualMotorDict)

				break
			elif(lineLenght-lineCounter == 0):
				#'Name' : 'zName Not found for Part Number',

				individualMotorDict = {'Name' : motorMechanicaName[x],
										'Assembly':assembly[x],
										'Part':secondPartNumberList[x],
										'Manufacturer':manufacturer[x],
										'AssemblyNumber': assemblyNumber[x],
										'Model':model[x],
										'Description':description[x],
										'QTY': quantity[x],
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

			if motor['QTY'] > 1:
				f.write("Quantity: ")
				f.write(str(motor['QTY']))
				f.write("\n")


			if "empty String" in motor['Name']:
				f.write("Assembly: ")
				f.write(motor['Assembly'])
				f.write("\n")
				f.write("Assembly PDF: ")
				f.write(motor['AssemblyNumber'])
				f.write("\n")
				f.write("Manufacturer: ")
				f.write(motor['Manufacturer'])
				f.write("\n")

			f.write("Part Number: ")
			f.write(motor['Part'])
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

			f.write("Assembly: ")
			f.write(motor['Assembly'])
			f.write("\n")

			f.write("Quantity: ")
			f.write(str(motor['QTY']))
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
	easygui.msgbox('After OK, the program will load database to RAM. This can take a minute.')
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
	"""
	['PRODCODE', 'CATINDEX', 'PARTNO', 'DESCRIPT', 'MANUFACTER', 'MODELNO', 'DRAWSIZE', 'DRAWINGNO', 'REVLEVEL', 'COST', 'POUNIT', 'PORATIO', 'UNIT',
	'ONHAND', 'ONORDER', 'ONDEMAND', 'WIPQTY', 'MINQTY', 'MAXQTY', 'LTIME', 'PREVQTY', 'SALEPRICE', 'AVAIL', 'ALTPARTNO', 'INVAREA1', 'INVAREA2', 'INVAREA3',
	'INVAREA4', 'LOCATE', 'PART_ASSY', 'ORDQTY', 'ORDMULT', 'STDCOST', 'MFG2', 'MFG3', 'CLASS', 'USAGE', 'MODELNO2', 'MODELNO3', 'IMAGE_FILE', 'NEWUSAGE', 'AREA2QTY',
	'AREA3QTY', 'AREA4QTY', 'AREA5QTY', 'AREA6QTY', 'SALESMAN', 'COMMISS', 'STARTDATE', 'WEIGHT', 'INVAREA5', 'INVAREA6', 'INVAREA7', 'INVAREA8', 'INVTOT', 'LOCATE2', 'LOCATE3', 'LOCATE4',
	'LOCATE5', 'LOCATE6', 'DIVISION', 'ALTPART1', 'ALTPART2', 'ALTPART3', 'ALTPART4', 'ALTPART5', 'ALTPART6', 'LASTPHYDAT', 'LASTQTY1', 'LASTQTY2', 'LASTQTY3', 'LASTQTY4', 'LASTQTY5', 'LASTQTY6',
	'LASTQTYWIP', 'LOGFILE', 'LASTPOCOST', 'MFG4', 'MFG5', 'MFG6', 'MFG7', 'MFG8', 'MFG9', 'MODELNO4', 'MODELNO5', 'MODELNO6', 'MODELNO7', 'MODELNO8', 'MODELNO9', 'DACCT1', 'STDLABCOST',
	'AVELABCOST', 'LPOLABCOST', 'LASTVENAME', 'LASTVENDID', 'FLOORSTK', 'SHELFLIFE', 'LICENSOR', 'ROYALRATE', 'PARTTYPE', 'VALUE', 'TOLERANCE', 'RATING', 'PACKTYPE', 'SCHEMATIC', 'FOOTPRINT', 'CACCT1',
	'BUYER', 'WEBITEM', 'PRICEKEY', 'SERIALITEM', 'TAGNUMBER', 'QBPARTID', 'QBASSET', 'QBINCOME', 'QBCOGSOLD', 'STDOUTCOST', 'AVEOUTCOST', 'LPOOUTCOST', 'OUTSOURCE', 'RELFACT', 'REOCCURING', 'OBSOLETE',
	'OBSDATE', 'SUBSONLY', 'AREA7QTY', 'AREA8QTY', 'AREA9QTY', 'AREA10QTY', 'AREA11QTY', 'INVAREA9', 'INVAREA10', 'INVAREA11', 'INVAREA12', 'LASTQTY7', 'LASTQTY8', 'LASTQTY9', 'LASTQTY10', 'LASTQTY11',
	'LOCATE7', 'LOCATE8', 'LOCATE9', 'LOCATE10', 'LOCATE11', 'ROHS', 'BOMSTATUS', 'SHIPPABLE', 'BLOWTHRU', 'AREA12QTY', 'AREA13QTY', 'AREA14QTY', 'AREA15QTY', 'AREA16QTY', 'AREA17QTY', 'AREA18QTY', 'AREA19QTY',
	'AREA20QTY', 'INVAREA13', 'INVAREA14', 'INVAREA15', 'INVAREA16', 'INVAREA17', 'INVAREA18', 'INVAREA19', 'INVAREA20', 'INVAREA21', 'LASTQTY12', 'LASTQTY13', 'LASTQTY14', 'LASTQTY15', 'LASTQTY16',
	'LASTQTY17', 'LASTQTY18', 'LASTQTY19', 'LASTQTY20', 'LOCATE12', 'LOCATE13', 'LOCATE14', 'LOCATE15', 'LOCATE16', 'LOCATE17', 'LOCATE18', 'LOCATE19', 'LOCATE20', 'COGSACCT', 'RECAREA', 'TARIFFCODE', 'NCNR',
	'SAFEWEEKS', 'MINQTYLOCK', 'WEEKSOFINV', 'COO', 'ECCN', 'AUX1', 'AUX2', 'ID1', 'WOI_QTY']
	"""
	#print('BOM TABLE HEADERS')
	#print (BOM_TABLE.field_names)
	"""
	['BOMNO', 'BOMDESCRI', 'PRODCODE', 'CATINDEX', 'PARTNO', 'ITEMNO', 'QTY', 'PART_ASSY', 'REFDES', 'REFDES2', 'REFDES3', 'REFDES4', 'ALTPART1', 'ALTPART2',
	'ALTPART3', 'ALTPART4', 'ALTPART5', 'ALTPART6', 'REFDESMEMO', 'KEYID', 'STAGEBIN', 'PARTDESC']
	"""


	while True:
		bomNumber = easygui.enterbox(msg="Enter the BOM you want to query.")

		#print (bomNumber[0])
		#print (type(bomNumber))
		if bomNumber[0] == "e":
			bomNumber = bomNumber.replace("e","E")

		if not bomNumber:
			sys.exit()

		#Complete list of all the parts in whole BOM. This is in record format
		recordsList = returnALLRecords(bomNumber, BOM_TABLE)

		if recordsList:
			break
		else:
			if not easygui.ynbox(msg="No BOM found with the number: "+bomNumber+"\n\nTry again?"):
				sys.exit()

	writePartNumbers(recordsList)
	motorsList = generateMotorList(recordsList,PART_TABLE)
	writeToOutput(motorsList[0],motorsList[1],getBomName(bomNumber,BOM_TABLE))
