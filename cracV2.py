import xlwt
from datetime import datetime
import os
import time 

class Unit():
	def __init__(self, name="unknown", temperature=0, humitty = 0):
		self.name = name
		self.temperature = temperature
		self.humitty = humitty
		
	def sys_health(self):
		if self.temperature > 85:
			print "Temperature is out of tolerance!"
		else:
			print "Temperature is within tolerance!"
		if self.humitty > 60:			
			print "Humitty is out of tolerance!"
		else:
			print "Humitty is within tolerance!"
			
name = raw_input('Enter your name: ')
shift = raw_input('Current shift: ')



wb = xlwt.Workbook()
ws = wb.add_sheet("A test sheet") 																			#this is to create the sheet but I need to change this to something relevant later
ws.write( 0, 1, 'Date: %s' % time.strftime("%x"))
ws.write( 0, 3, 'Location: LAX1')
ws.write( 0, 5, 'Performed By: %s' % name)
ws.write( 0, 9, 'Shift: %s' % shift)
saveAs = raw_input("Save file as: ") 																		#Save mechanism
saveAs = saveAs + ".xls"
print "Saving as: %s" % saveAs
wb.save(saveAs) 

#############################################################################################################
endOfday = False
while endOfday == False:

	unitIdRow = 1 																									#This is the first Row/Column to be written to for the crac unit's ID.
	unitIdColumn = 1
	unitTempRow = 3 																								#This is the first Row/Column to be written to for the first units temperature.
	unitTempColumn = 1
	unitHumittyRow = 3 																								#This is the first Row/Column to be written to for the first units humitty.
	unitHumittyColumn = 2
	tempDesignationRow = 2																							#This is the first Row/Column to be written to for the Designation of temperature column's.
	tempDesignationColumn = 1
	humittyDesignationRow = 2																						#This is the first Row/Column to be written to for the Designation of humitty column's.
	humittyDesignationColumn = 2
	timeDesignationRow = 3																							#This is the first Row/Column to be written to for the Designation of the time of day.
	timeDesignationColumn = 0

	validEntry = False

	while validEntry == False:
		hour = int(raw_input('Enter current hour (8, 9, 10, 11, 12, 1, 2, 3): '))
		if hour == 8:
			unitTempRow = unitTempRow + 0
			unitHumittyRow = unitHumittyRow + 0
			validEntry = True
		elif hour == 9:
			unitTempRow = unitTempRow + 1
			unitHumittyRow = unitHumittyRow + 1
			validEntry = True
		elif hour == 10:
			unitTempRow = unitTempRow + 2
			unitHumittyRow = unitHumittyRow + 2
			validEntry = True
		elif hour == 11:
			unitTempRow = unitTempRow + 3
			unitHumittyRow = unitHumittyRow + 3
			validEntry = True
		elif hour == 12:
			unitTempRow = unitTempRow + 4
			unitHumittyRow = unitHumittyRow + 4
			validEntry = True
		elif hour == 1:
			unitTempRow = unitTempRow + 5
			unitHumittyRow = unitHumittyRow + 5
			validEntry = True
		elif hour == 2:
			unitTempRow = unitTempRow + 6
			unitHumittyRow = unitHumittyRow + 6
			validEntry = True
		elif hour == 3:
			unitTempRow = unitTempRow + 7
			unitHumittyRow = unitHumittyRow + 7
			validEntry = True
		else: 
			print 'error'
			validEntry = False

	activeUnits = ['CRAC 135-1', 'CRAC 135-2', 'CRAC 135-3', 'CRAC 135-4', 'CRAC 135-5', 'CRAC 135-6', 'CRAC 135-7', 'CRAC 135-8', 'CRAC 135-9', 'CRAC 135-10', 'CRAC 135-11']			#Need to have a way of altering this list 

	timeFirstShift = ['8:00 AM', '9:00 AM', '10:00 AM', '11:00 AM', '12:00 PM', '1:00 PM', '2:00 PM', '3:00 PM']

	for i in activeUnits: 																							# This loop is for printing the headers that should not be reprinted.
		ws.write( tempDesignationRow, tempDesignationColumn, "Temp.")
		tempDesignationColumn = tempDesignationColumn + 2
		
		ws.write( humittyDesignationRow, humittyDesignationColumn, "Humitty")
		humittyDesignationColumn = humittyDesignationColumn + 2
		
	for i in timeFirstShift: 																						#print out time's 
		ws.write(timeDesignationRow, timeDesignationColumn, i)
		timeDesignationRow = timeDesignationRow + 1
		

	for i in activeUnits: 																							#This loop is for getting temp and humitty from the user and writing them to spreadsheet.
		print "__________________________________________________________"
		
		print "Current unit %s " % i
		ws.write( unitIdRow, unitIdColumn, i)
		unitIdColumn = unitIdColumn + 2
		
		currentUnitTemp = int(raw_input("Enter the current temperature for unit %s: " % i))
		ws.write( unitTempRow, unitTempColumn, currentUnitTemp)
		unitTempColumn = unitTempColumn + 2

		
		currentUnitHumitty = int(raw_input("Enter the current humitty for unit %s: " % i))
		ws.write( unitHumittyRow, unitHumittyColumn, currentUnitHumitty)
		unitHumittyColumn = unitHumittyColumn + 2

		
		i = Unit(i, currentUnitTemp, currentUnitHumitty)
		i.sys_health()
		
	#############################################################################################################
		
	unitIdRow = 11 																									#This is the first Row/Column to be written to for the crac unit's ID.
	unitIdColumn = 1
	unitTempRow = 13 																								#This is the first Row/Column to be written to for the first units temperature.
	unitTempColumn = 1
	unitHumittyRow = 13 																								#This is the first Row/Column to be written to for the first units humitty.
	unitHumittyColumn = 2
	tempDesignationRow = 12																							#This is the first Row/Column to be written to for the Designation of temperature column's.
	tempDesignationColumn = 1
	humittyDesignationRow = 12																						#This is the first Row/Column to be written to for the Designation of humitty column's.
	humittyDesignationColumn = 2
	timeDesignationRow = 13																							#This is the first Row/Column to be written to for the Designation of the time of day.
	timeDesignationColumn = 0

	if hour == 8:
		unitTempRow = unitTempRow + 0
		unitHumittyRow = unitHumittyRow + 0
	elif hour == 9:
		unitTempRow = unitTempRow + 1
		unitHumittyRow = unitHumittyRow + 1
	elif hour == 10:
		unitTempRow = unitTempRow + 2
		unitHumittyRow = unitHumittyRow + 2
	elif hour == 11:
		unitTempRow = unitTempRow + 3
		unitHumittyRow = unitHumittyRow + 3
	elif hour == 12:
		unitTempRow = unitTempRow + 4
		unitHumittyRow = unitHumittyRow + 4
	elif hour == 1:
		unitTempRow = unitTempRow + 5
		unitHumittyRow = unitHumittyRow + 5
	elif hour == 2:
		unitTempRow = unitTempRow + 6
		unitHumittyRow = unitHumittyRow + 6
	elif hour == 3:
		unitTempRow = unitTempRow + 7
		unitHumittyRow = unitHumittyRow + 7
	else: 
		print 'error'	
		
	activeUnits2 = [ 'CRAC 135-12', 'CRAC 135-13', 'CRAC 135-14','CRAC 130-15', 'CRAC 130-16', 'CRAC 130-17', 'CRAC 130-18', 'CRAC 130-19', 'CRAC 130-20', 'CRAC 130-21', 'CRAC 130-22']


	for i in activeUnits2: 																							# This loop is for printing the headers that should not be reprinted.
		ws.write( tempDesignationRow, tempDesignationColumn, "Temp.")
		tempDesignationColumn = tempDesignationColumn + 2
		
		ws.write( humittyDesignationRow, humittyDesignationColumn, "Humitty")
		humittyDesignationColumn = humittyDesignationColumn + 2
		
	for i in timeFirstShift: 																						#print out time's 
		ws.write(timeDesignationRow, timeDesignationColumn, i)
		timeDesignationRow = timeDesignationRow + 1
		
	for i in activeUnits2: 																							#This loop is for getting temp and humitty from the user and writing them to spreadsheet.
		print "__________________________________________________________"
		
		print "Current unit %s " % i
		ws.write( unitIdRow, unitIdColumn, i)
		unitIdColumn = unitIdColumn + 2
		
		currentUnitTemp = int(raw_input("Enter the current temperature for unit %s: " % i))
		ws.write( unitTempRow, unitTempColumn, currentUnitTemp)
		unitTempColumn = unitTempColumn + 2

		
		currentUnitHumitty = int(raw_input("Enter the current humitty for unit %s: " % i))
		ws.write( unitHumittyRow, unitHumittyColumn, currentUnitHumitty)
		unitHumittyColumn = unitHumittyColumn + 2

		
		i = Unit(i, currentUnitTemp, currentUnitHumitty)
		i.sys_health()
		
	#############################################################################################################
		
	unitIdRow = 21 																									#This is the first Row/Column to be written to for the crac unit's ID.
	unitIdColumn = 1
	unitTempRow = 23 																								#This is the first Row/Column to be written to for the first units temperature.
	unitTempColumn = 1
	unitHumittyRow = 23 																								#This is the first Row/Column to be written to for the first units humitty.
	unitHumittyColumn = 2
	tempDesignationRow = 22																							#This is the first Row/Column to be written to for the Designation of temperature column's.
	tempDesignationColumn = 1
	humittyDesignationRow = 22																						#This is the first Row/Column to be written to for the Designation of humitty column's.
	humittyDesignationColumn = 2
	timeDesignationRow = 23																							#This is the first Row/Column to be written to for the Designation of the time of day.
	timeDesignationColumn = 0

	if hour == 8:
		unitTempRow = unitTempRow + 0
		unitHumittyRow = unitHumittyRow + 0
	elif hour == 9:
		unitTempRow = unitTempRow + 1
		unitHumittyRow = unitHumittyRow + 1
	elif hour == 10:
		unitTempRow = unitTempRow + 2
		unitHumittyRow = unitHumittyRow + 2
	elif hour == 11:
		unitTempRow = unitTempRow + 3
		unitHumittyRow = unitHumittyRow + 3
	elif hour == 12:
		unitTempRow = unitTempRow + 4
		unitHumittyRow = unitHumittyRow + 4
	elif hour == 1:
		unitTempRow = unitTempRow + 5
		unitHumittyRow = unitHumittyRow + 5
	elif hour == 2:
		unitTempRow = unitTempRow + 6
		unitHumittyRow = unitHumittyRow + 6
	elif hour == 3:
		unitTempRow = unitTempRow + 7
		unitHumittyRow = unitHumittyRow + 7
	else: 
		print 'error'	
		
	activeUnits3 = ['CRAC 130-23', 'CRAC 130-24', 'CRAC 130-25', 'CRAC 130-26', 'CRAC 240-02', 'CRAC 240-03', 'CRAC 240-04', 'CRAC 240-05', 'CRAC 240-06', 'CRAC 240-07', 'CRAC 240-08']


	for i in activeUnits3: 																							# This loop is for printing the headers that should not be reprinted.
		ws.write( tempDesignationRow, tempDesignationColumn, "Temp.")
		tempDesignationColumn = tempDesignationColumn + 2
		
		ws.write( humittyDesignationRow, humittyDesignationColumn, "Humitty")
		humittyDesignationColumn = humittyDesignationColumn + 2
		
	for i in timeFirstShift: 																						#print out time's 
		ws.write(timeDesignationRow, timeDesignationColumn, i)
		timeDesignationRow = timeDesignationRow + 1
		
	for i in activeUnits3: 																							#This loop is for getting temp and humitty from the user and writing them to spreadsheet.
		print "__________________________________________________________"
		
		print "Current unit %s " % i
		ws.write( unitIdRow, unitIdColumn, i)
		unitIdColumn = unitIdColumn + 2
		
		currentUnitTemp = int(raw_input("Enter the current temperature for unit %s: " % i))
		ws.write( unitTempRow, unitTempColumn, currentUnitTemp)
		unitTempColumn = unitTempColumn + 2

		
		currentUnitHumitty = int(raw_input("Enter the current humitty for unit %s: " % i))
		ws.write( unitHumittyRow, unitHumittyColumn, currentUnitHumitty)
		unitHumittyColumn = unitHumittyColumn + 2

		
		i = Unit(i, currentUnitTemp, currentUnitHumitty)
		i.sys_health()
		
	#############################################################################################################

	unitIdRow = 31 																									#This is the first Row/Column to be written to for the crac unit's ID.
	unitIdColumn = 1
	unitTempRow = 33 																								#This is the first Row/Column to be written to for the first units temperature.
	unitTempColumn = 1
	unitHumittyRow = 33 																								#This is the first Row/Column to be written to for the first units humitty.
	unitHumittyColumn = 2
	tempDesignationRow = 32																							#This is the first Row/Column to be written to for the Designation of temperature column's.
	tempDesignationColumn = 1
	humittyDesignationRow = 32																						#This is the first Row/Column to be written to for the Designation of humitty column's.
	humittyDesignationColumn = 2
	timeDesignationRow = 33																							#This is the first Row/Column to be written to for the Designation of the time of day.
	timeDesignationColumn = 0

	if hour == 8:
		unitTempRow = unitTempRow + 0
		unitHumittyRow = unitHumittyRow + 0
	elif hour == 9:
		unitTempRow = unitTempRow + 1
		unitHumittyRow = unitHumittyRow + 1
	elif hour == 10:
		unitTempRow = unitTempRow + 2
		unitHumittyRow = unitHumittyRow + 2
	elif hour == 11:
		unitTempRow = unitTempRow + 3
		unitHumittyRow = unitHumittyRow + 3
	elif hour == 12:
		unitTempRow = unitTempRow + 4
		unitHumittyRow = unitHumittyRow + 4
	elif hour == 1:
		unitTempRow = unitTempRow + 5
		unitHumittyRow = unitHumittyRow + 5
	elif hour == 2:
		unitTempRow = unitTempRow + 6
		unitHumittyRow = unitHumittyRow + 6
	elif hour == 3:
		unitTempRow = unitTempRow + 7
		unitHumittyRow = unitHumittyRow + 7
	else: 
		print 'error'	
		
	activeUnits4 = ['CRAC 240-09', 'CRAC 240-12', 'CRAC 240-13', 'CRAC 240-14']


	for i in activeUnits4: 																							# This loop is for printing the headers that should not be reprinted.
		ws.write( tempDesignationRow, tempDesignationColumn, "Temp.")
		tempDesignationColumn = tempDesignationColumn + 2
		
		ws.write( humittyDesignationRow, humittyDesignationColumn, "Humitty")
		humittyDesignationColumn = humittyDesignationColumn + 2
		
	for i in timeFirstShift: 																						#print out time's 
		ws.write(timeDesignationRow, timeDesignationColumn, i)
		timeDesignationRow = timeDesignationRow + 1
		
	for i in activeUnits4: 																							#This loop is for getting temp and humitty from the user and writing them to spreadsheet.
		print "__________________________________________________________"
		
		print "Current unit %s " % i
		ws.write( unitIdRow, unitIdColumn, i)
		unitIdColumn = unitIdColumn + 2
		
		currentUnitTemp = int(raw_input("Enter the current temperature for unit %s: " % i))
		ws.write( unitTempRow, unitTempColumn, currentUnitTemp)
		unitTempColumn = unitTempColumn + 2

		
		currentUnitHumitty = int(raw_input("Enter the current humitty for unit %s: " % i))
		ws.write( unitHumittyRow, unitHumittyColumn, currentUnitHumitty)
		unitHumittyColumn = unitHumittyColumn + 2

		
		i = Unit(i, currentUnitTemp, currentUnitHumitty)
		i.sys_health()
		
	########################################################################################################################################################
		
	activeUps = ['UPS 130-1', 'UPS 130-2', 'UPS 135-1', 'UPS 135-2', 'UPS 135-3', 'UPS 240-1', 'UPS 240-2']

	upsStatusDesignationRow = 32
	upsStatusDesignationColumn = 9
	upsIdDesignationRow = 31
	upsIdDesignationColumn = 9
	upsStatusRow = 33
	upsStatusColumn = 9

	if hour == 8:
		upsStatusRow = upsStatusRow + 0
	elif hour == 9:
		upsStatusRow = upsStatusRow + 1
	elif hour == 10:
		upsStatusRow = upsStatusRow + 2
	elif hour == 11:
		upsStatusRow = upsStatusRow + 3
	elif hour == 12:
		upsStatusRow = upsStatusRow + 4
	elif hour == 1:
		upsStatusRow = upsStatusRow + 5
	elif hour == 2:
		upsStatusRow = upsStatusRow + 6
	elif hour == 3:
		upsStatusRow = upsStatusRow + 7
	else: 
		print 'error'	

	for i in activeUps: 																							# This loop is for printing the headers that should not be reprinted.
		ws.write( upsStatusDesignationRow, upsStatusDesignationColumn, "System Normal")
		upsStatusDesignationColumn = upsStatusDesignationColumn + 2
		
	for i in activeUps: 																							#This loop is for getting temp and humitty from the user and writing them to spreadsheet.
		print "__________________________________________________________"
		
		print "Current UPS %s " % i
		ws.write( upsIdDesignationRow, upsIdDesignationColumn, i)
		upsIdDesignationColumn = upsIdDesignationColumn + 2
		currentUps = i

		onOff = raw_input("Is %s fully operational?(y/n) " % currentUps)
		if onOff == 'y':
			currentUpsStatus = 'True'
		elif onOff == 'n':
			currentUpsStatus = 'False'
		else:
			print "Error"
			currentUpsStatus = "Error"

		ws.write( upsStatusRow, upsStatusColumn, currentUpsStatus)
		upsStatusColumn = upsStatusColumn + 2
		
	print "__________________________________________________________"
	print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
		
	#saveAs = raw_input("Save file as: ") 																		#Save mechanism
	#saveAs = saveAs + ".xls"
	print "Saving as: %s" % saveAs
	wb.save(saveAs) 																							#May want to automate this at some point for easy of use and standardization reasons.

	print "Opening %s" % saveAs																					#Open the newly created spreadsheet
	os.system("start %s" % saveAs)
	print "Success!"
	
	validEntry = False
	while validEntry == False:
		endOfday = raw_input('End of day?(y,n): ')
		if endOfday == 'y':
			endOfday = True
			validEntry = True
		elif endOfday == 'Y':
			endOfday = True
			validEntry = True
		elif endOfday == 'n':
			endOfday = False
			validEntry = True
		elif endOfday == 'N':
			endOfday = False
			validEntry = True
		else:
			validEntry = False