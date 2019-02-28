import xlwt
from datetime import datetime
import os
import time 
import sys 

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

validEntry = False
while validEntry == False:
	shift = int(raw_input('Current shift(1-3): '))
	if shift == 1:
		timeFirstShift = ['8:00 AM', '9:00 AM', '10:00 AM', '11:00 AM', '12:00 PM', '1:00 PM', '2:00 PM', '3:00 PM']
		hourList = ['8am', '9am', '10am', '11am', '12pm', '1pm', '2pm', '3pm']
		validEntry = True
	elif shift == 2:
		timeFirstShift = ['4:00 PM', '5:00 PM', '6:00 PM', '7:00 PM', '8:00 PM', '9:00 PM', '10:00 PM', '11:00 PM']
		hourList = ['4pm', '5pm', '6pm', '7pm', '8pm', '9pm', '10pm', '11pm']
		validEntry = True
	elif shift == 3:
		timeFirstShift = ['12:00 AM', '1:00 AM', '2:00 AM', '3:00 AM', '4:00 AM', '5:00 AM', '6:00 AM', '7:00 AM']
		hourList = ['12am', '1am', '2am', '3am', '4am', '5am', '6am', '7am']
		validEntry = True
	else:
		print "Error! %s is not a valid shift!" % shift

saveAs = raw_input("Save file as: ")
saveAs = saveAs + ".xls"
print "Saving as: %s" % saveAs																		
wb = xlwt.Workbook()
ws = wb.add_sheet('WalkThrough') 																			
ws.write( 0, 1, 'Date: %s' % time.strftime("%x"))
ws.write( 0, 3, 'Location: LAX1')
ws.write( 0, 5, 'Performed By: %s' % name)
ws.write( 0, 9, 'Shift: %s' % shift)


while True:
	try:
		wb.save(saveAs)
		break
	except IOError:
		print "Cannot write to open sheet!" 
		raw_input('Close the sheet then press any key to continue.')
		
unitIdRow = 1 																									
unitIdColumn = 1
tempDesignationRow = 2																							
tempDesignationColumn = 1
humittyDesignationRow = 2																						
humittyDesignationColumn = 2
timeDesignationRow = 3																							
timeDesignationColumn = 0

activeUnits = ['CRAC 135-1', 'CRAC 135-2', 'CRAC 135-3', 'CRAC 135-4', 'CRAC 135-5', 'CRAC 135-6', 'CRAC 135-7', 'CRAC 135-8', 'CRAC 135-9', 'CRAC 135-10', 'CRAC 135-11']

for i in activeUnits: 																							# This loop is for printing the headers that should not be reprinted.
	ws.write( tempDesignationRow, tempDesignationColumn, "Temp.")
	tempDesignationColumn = tempDesignationColumn + 2
	ws.write( humittyDesignationRow, humittyDesignationColumn, "Humitty")
	humittyDesignationColumn = humittyDesignationColumn + 2
	ws.write( unitIdRow, unitIdColumn, i)
	unitIdColumn = unitIdColumn + 2
	
for i in timeFirstShift: 																						#print out time's 
	ws.write(timeDesignationRow, timeDesignationColumn, i)
	timeDesignationRow = timeDesignationRow + 1

unitIdRow = 11 																									#This is the first Row/Column to be written to for the crac unit's ID.
unitIdColumn = 1
tempDesignationRow = 12																							#This is the first Row/Column to be written to for the Designation of temperature column's.
tempDesignationColumn = 1
humittyDesignationRow = 12																						#This is the first Row/Column to be written to for the Designation of humitty column's.
humittyDesignationColumn = 2
timeDesignationRow = 13																							#This is the first Row/Column to be written to for the Designation of the time of day.
timeDesignationColumn = 0

activeUnits2 = [ 'CRAC 135-12', 'CRAC 135-13', 'CRAC 135-14','CRAC 130-15', 'CRAC 130-16', 'CRAC 130-17', 'CRAC 130-18', 'CRAC 130-19', 'CRAC 130-20', 'CRAC 130-21', 'CRAC 130-22']

for i in activeUnits2: 																							# This loop is for printing the headers that should not be reprinted.
	ws.write( tempDesignationRow, tempDesignationColumn, "Temp.")
	tempDesignationColumn = tempDesignationColumn + 2
	ws.write( humittyDesignationRow, humittyDesignationColumn, "Humitty")
	humittyDesignationColumn = humittyDesignationColumn + 2
	ws.write( unitIdRow, unitIdColumn, i)
	unitIdColumn = unitIdColumn + 2

for i in timeFirstShift: 																						#print out time's 
	ws.write(timeDesignationRow, timeDesignationColumn, i)
	timeDesignationRow = timeDesignationRow + 1
	
unitIdRow = 21 																									#This is the first Row/Column to be written to for the crac unit's ID.
unitIdColumn = 1
tempDesignationRow = 22																							#This is the first Row/Column to be written to for the Designation of temperature column's.
tempDesignationColumn = 1
humittyDesignationRow = 22																						#This is the first Row/Column to be written to for the Designation of humitty column's.
humittyDesignationColumn = 2
timeDesignationRow = 23																							#This is the first Row/Column to be written to for the Designation of the time of day.
timeDesignationColumn = 0

activeUnits3 = ['CRAC 130-23', 'CRAC 130-24', 'CRAC 130-25', 'CRAC 130-26', 'CRAC 240-02', 'CRAC 240-03', 'CRAC 240-04', 'CRAC 240-05', 'CRAC 240-06', 'CRAC 240-07', 'CRAC 240-08']

for i in activeUnits3: 																							# This loop is for printing the headers that should not be reprinted.
	ws.write( tempDesignationRow, tempDesignationColumn, "Temp.")
	tempDesignationColumn = tempDesignationColumn + 2
	ws.write( humittyDesignationRow, humittyDesignationColumn, "Humitty")
	humittyDesignationColumn = humittyDesignationColumn + 2
	ws.write( unitIdRow, unitIdColumn, i)
	unitIdColumn = unitIdColumn + 2

for i in timeFirstShift: 																						#print out time's 
	ws.write(timeDesignationRow, timeDesignationColumn, i)
	timeDesignationRow = timeDesignationRow + 1

unitIdRow = 31 																									#This is the first Row/Column to be written to for the crac unit's ID.
unitIdColumn = 1
tempDesignationRow = 32																							#This is the first Row/Column to be written to for the Designation of temperature column's.
tempDesignationColumn = 1
humittyDesignationRow = 32																						#This is the first Row/Column to be written to for the Designation of humitty column's.
humittyDesignationColumn = 2
timeDesignationRow = 33																							#This is the first Row/Column to be written to for the Designation of the time of day.
timeDesignationColumn = 0

activeUnits4 = ['CRAC 240-09', 'CRAC 240-12', 'CRAC 240-13', 'CRAC 240-14']

for i in activeUnits4: 																							# This loop is for printing the headers that should not be reprinted.
	ws.write( tempDesignationRow, tempDesignationColumn, "Temp.")
	tempDesignationColumn = tempDesignationColumn + 2
	ws.write( humittyDesignationRow, humittyDesignationColumn, "Humitty")
	humittyDesignationColumn = humittyDesignationColumn + 2
	ws.write( unitIdRow, unitIdColumn, i)
	unitIdColumn = unitIdColumn + 2
	
for i in timeFirstShift: 																						#print out time's 
	ws.write(timeDesignationRow, timeDesignationColumn, i)
	timeDesignationRow = timeDesignationRow + 1

activeUps = ['UPS 130-1', 'UPS 130-2', 'UPS 135-1', 'UPS 135-2', 'UPS 135-3', 'UPS 240-1', 'UPS 240-2']

upsStatusDesignationRow = 32
upsStatusDesignationColumn = 9
upsIdDesignationRow = 31
upsIdDesignationColumn = 9

for i in activeUps: 																							# This loop is for printing the headers that should not be reprinted.
	ws.write( upsStatusDesignationRow, upsStatusDesignationColumn, "System Normal")
	upsStatusDesignationColumn = upsStatusDesignationColumn + 2
	ws.write( upsIdDesignationRow, upsIdDesignationColumn, i)
	upsIdDesignationColumn = upsIdDesignationColumn + 2
	
notes = ''
notesDesignationRow = 43
notesDesignationColumn = 1
ws.write( notesDesignationRow, notesDesignationColumn, "Notes: ")
notesDesignationRow = 44
notesDesignationColumn = 2
outOfToleranceUnits = []
notes = "Units out of tolerance: %s " % outOfToleranceUnits

#############################################################################################################
endOfday = False
while endOfday == False:

	unitTempRow = 3 																								#This is the first Row/Column to be written to for the first units temperature.
	unitTempColumn = 1
	unitHumittyRow = 3 																								#This is the first Row/Column to be written to for the first units humitty.
	unitHumittyColumn = 2

	validEntry = False

	while validEntry == False:
		hour = raw_input('Enter current hour %s: ' % hourList)
		if hour == '8am' or '4pm' or '12am':
			unitTempRow = unitTempRow + 0
			unitHumittyRow = unitHumittyRow + 0
			validEntry = True
		elif hour == '9am' or '5pm' or '1am':
			unitTempRow = unitTempRow + 1
			unitHumittyRow = unitHumittyRow + 1
			validEntry = True
		elif hour == '10am' or '6pm' or '2am':
			unitTempRow = unitTempRow + 2
			unitHumittyRow = unitHumittyRow + 2
			validEntry = True
		elif hour == '11am' or '7pm' or '3am':
			unitTempRow = unitTempRow + 3
			unitHumittyRow = unitHumittyRow + 3
			validEntry = True
		elif hour == '12pm' or '8pm' or '4am':
			unitTempRow = unitTempRow + 4
			unitHumittyRow = unitHumittyRow + 4
			validEntry = True
		elif hour == '1pm' or '9pm' or '5am':
			unitTempRow = unitTempRow + 5
			unitHumittyRow = unitHumittyRow + 5
			validEntry = True
		elif hour == '2pm' or '10pm' or '6am':
			unitTempRow = unitTempRow + 6
			unitHumittyRow = unitHumittyRow + 6
			validEntry = True
		elif hour == '3pm' or '11pm' or '7am':
			unitTempRow = unitTempRow + 7
			unitHumittyRow = unitHumittyRow + 7
			validEntry = True
		else: 
			print 'error'
			validEntry = False


	for i in activeUnits: 																							#This loop is for getting temp and humitty from the user and writing them to spreadsheet.
		print "__________________________________________________________"
		
		print "Current unit %s " % i

		while True:
			try:
				currentUnitTemp = int(raw_input("Enter the current temperature for unit %s: " % i))
				break
			except ValueError:
				print "Invalid Entry (Must be a number!)"	
			except EOFError:
				print "Quitting!"
				exit()
			except KeyboardInterrupt:
				print "Quitting!"
				exit()
				
		ws.write( unitTempRow, unitTempColumn, currentUnitTemp)
		unitTempColumn = unitTempColumn + 2

		while True:
			try:
				currentUnitHumitty = int(raw_input("Enter the current humitty for unit %s: " % i))
				break
			except ValueError:
				print "Invalid Entry (Must be a number!)"
			except EOFError:
				print "Quitting!"
				exit()
			except KeyboardInterrupt:
				print "Quitting!"
				exit()
				
		ws.write( unitHumittyRow, unitHumittyColumn, currentUnitHumitty)
		unitHumittyColumn = unitHumittyColumn + 2

		if currentUnitTemp > 80 or currentUnitHumitty > 50:
			outOfToleranceUnits.append([i])
			
		i = Unit(i, currentUnitTemp, currentUnitHumitty)
		i.sys_health()
		print outOfToleranceUnits
	#############################################################################################################
		

	unitTempRow = 13 																								#This is the first Row/Column to be written to for the first units temperature.
	unitTempColumn = 1
	unitHumittyRow = 13 																								#This is the first Row/Column to be written to for the first units humitty.
	unitHumittyColumn = 2


	if hour == '8am' or '4pm' or '12am':
		unitTempRow = unitTempRow + 0
		unitHumittyRow = unitHumittyRow + 0
	elif hour == '9am' or '5pm' or '1am':
		unitTempRow = unitTempRow + 1
		unitHumittyRow = unitHumittyRow + 1
	elif hour == '10am' or '6pm' or '2am':
		unitTempRow = unitTempRow + 2
		unitHumittyRow = unitHumittyRow + 2
	elif hour == '11am' or '7pm' or '3am':
		unitTempRow = unitTempRow + 3
		unitHumittyRow = unitHumittyRow + 3
	elif hour == '12pm' or '8pm' or '4am':
		unitTempRow = unitTempRow + 4
		unitHumittyRow = unitHumittyRow + 4
	elif hour == '1pm' or '9pm' or '5am':
		unitTempRow = unitTempRow + 5
		unitHumittyRow = unitHumittyRow + 5
	elif hour == '2pm' or '10pm' or '6am':
		unitTempRow = unitTempRow + 6
		unitHumittyRow = unitHumittyRow + 6
	elif hour == '3pm' or '11pm' or '7am':
		unitTempRow = unitTempRow + 7
		unitHumittyRow = unitHumittyRow + 7
	else: 
		print 'error'	
		

		
	for i in activeUnits2: 																							#This loop is for getting temp and humitty from the user and writing them to spreadsheet.
		print "__________________________________________________________"
		
		print "Current unit %s " % i
		
		while True:
			try:
				currentUnitTemp = int(raw_input("Enter the current temperature for unit %s: " % i))
				break
			except ValueError:
				print "Invalid Entry (Must be a number!)"
			except EOFError:
				print "Quitting!"
				exit()
			except KeyboardInterrupt:
				print "Quitting!"
				exit()
				
		ws.write( unitTempRow, unitTempColumn, currentUnitTemp)
		unitTempColumn = unitTempColumn + 2
		
		while True:
			try:
				currentUnitHumitty = int(raw_input("Enter the current humitty for unit %s: " % i))
				break
			except ValueError:
				print "Invalid Entry (Must be a number!)"
			except EOFError:
				print "Quitting!"
				exit()
			except KeyboardInterrupt:
				print "Quitting!"
				exit()
				
		ws.write( unitHumittyRow, unitHumittyColumn, currentUnitHumitty)
		unitHumittyColumn = unitHumittyColumn + 2

		if currentUnitTemp > 80 or currentUnitHumitty > 50:
			outOfToleranceUnits.append([i])			
			
		i = Unit(i, currentUnitTemp, currentUnitHumitty)
		i.sys_health()
		
	#############################################################################################################
		
	unitTempRow = 23 																								#This is the first Row/Column to be written to for the first units temperature.
	unitTempColumn = 1
	unitHumittyRow = 23 																								#This is the first Row/Column to be written to for the first units humitty.
	unitHumittyColumn = 2

	if hour == '8am' or '4pm' or '12am':
		unitTempRow = unitTempRow + 0
		unitHumittyRow = unitHumittyRow + 0
	elif hour == '9am' or '5pm' or '1am':
		unitTempRow = unitTempRow + 1
		unitHumittyRow = unitHumittyRow + 1
	elif hour == '10am' or '6pm' or '2am':
		unitTempRow = unitTempRow + 2
		unitHumittyRow = unitHumittyRow + 2
	elif hour == '11am' or '7pm' or '3am':
		unitTempRow = unitTempRow + 3
		unitHumittyRow = unitHumittyRow + 3
	elif hour == '12pm' or '8pm' or '4am':
		unitTempRow = unitTempRow + 4
		unitHumittyRow = unitHumittyRow + 4
	elif hour == '1pm' or '9pm' or '5am':
		unitTempRow = unitTempRow + 5
		unitHumittyRow = unitHumittyRow + 5
	elif hour == '2pm' or '10pm' or '6am':
		unitTempRow = unitTempRow + 6
		unitHumittyRow = unitHumittyRow + 6
	elif hour == '3pm' or '11pm' or '7am':
		unitTempRow = unitTempRow + 7
		unitHumittyRow = unitHumittyRow + 7
	else: 
		print 'error'	
		
	for i in activeUnits3: 																							#This loop is for getting temp and humitty from the user and writing them to spreadsheet.
		print "__________________________________________________________"
		
		print "Current unit %s " % i
		
		while True:
			try:
				currentUnitTemp = int(raw_input("Enter the current temperature for unit %s: " % i))
				break
			except ValueError:
				print "Invalid Entry (Must be a number!)"
			except EOFError:
				print "Quitting!"
				exit()
			except KeyboardInterrupt:
				print "Quitting!"
				exit()
				
		ws.write( unitTempRow, unitTempColumn, currentUnitTemp)
		unitTempColumn = unitTempColumn + 2
		while True:
			try:
				currentUnitHumitty = int(raw_input("Enter the current humitty for unit %s: " % i))
				break
			except ValueError:
				print "Invalid Entry (Must be a number!)"
			except EOFError:
				print "Quitting!"
				exit()
			except KeyboardInterrupt:
				print "Quitting!"
				exit()
				
		ws.write( unitHumittyRow, unitHumittyColumn, currentUnitHumitty)
		unitHumittyColumn = unitHumittyColumn + 2
		
		if currentUnitTemp > 80 or currentUnitHumitty > 50:
			outOfToleranceUnits.append([i])
			
		i = Unit(i, currentUnitTemp, currentUnitHumitty)
		i.sys_health()
		
	#############################################################################################################

	unitTempRow = 33 																								#This is the first Row/Column to be written to for the first units temperature.
	unitTempColumn = 1
	unitHumittyRow = 33 																								#This is the first Row/Column to be written to for the first units humitty.
	unitHumittyColumn = 2

	if hour == '8am' or '4pm' or '12am':
		unitTempRow = unitTempRow + 0
		unitHumittyRow = unitHumittyRow + 0
	elif hour == '9am' or '5pm' or '1am':
		unitTempRow = unitTempRow + 1
		unitHumittyRow = unitHumittyRow + 1
	elif hour == '10am' or '6pm' or '2am':
		unitTempRow = unitTempRow + 2
		unitHumittyRow = unitHumittyRow + 2
	elif hour == '11am' or '7pm' or '3am':
		unitTempRow = unitTempRow + 3
		unitHumittyRow = unitHumittyRow + 3
	elif hour == '12pm' or '8pm' or '4am':
		unitTempRow = unitTempRow + 4
		unitHumittyRow = unitHumittyRow + 4
	elif hour == '1pm' or '9pm' or '5am':
		unitTempRow = unitTempRow + 5
		unitHumittyRow = unitHumittyRow + 5
	elif hour == '2pm' or '10pm' or '6am':
		unitTempRow = unitTempRow + 6
		unitHumittyRow = unitHumittyRow + 6
	elif hour == '3pm' or '11pm' or '7am':
		unitTempRow = unitTempRow + 7
		unitHumittyRow = unitHumittyRow + 7
	else: 
		print 'error'	
		
	for i in activeUnits4: 																							#This loop is for getting temp and humitty from the user and writing them to spreadsheet.
		print "__________________________________________________________"
		
		print "Current unit %s " % i
		
		while True:
			try:
				currentUnitTemp = int(raw_input("Enter the current temperature for unit %s: " % i))
				break
			except ValueError:
				print "Invalid Entry (Must be a number!)"
			except EOFError:
				print "Quitting!"
				exit()
			except KeyboardInterrupt:
				print "Quitting!"
				exit()
				
		ws.write( unitTempRow, unitTempColumn, currentUnitTemp)
		unitTempColumn = unitTempColumn + 2
		
		while True:
			try:
				currentUnitHumitty = int(raw_input("Enter the current humitty for unit %s: " % i))
				break
			except ValueError:
				print "Invalid Entry (Must be a number!)"
			except EOFError:
				print "Quitting!"
				exit()
			except KeyboardInterrupt:
				print "Quitting!"
				exit()
				
		ws.write( unitHumittyRow, unitHumittyColumn, currentUnitHumitty)
		unitHumittyColumn = unitHumittyColumn + 2
		
		if currentUnitTemp > 80 or currentUnitHumitty > 50:
			outOfToleranceUnits.append([i])
			
		i = Unit(i, currentUnitTemp, currentUnitHumitty)
		i.sys_health()
		
	########################################################################################################################################################
		
	upsStatusRow = 33
	upsStatusColumn = 9

	if hour == '8am' or '4pm' or '12am':
		upsStatusRow = upsStatusRow + 0
	elif hour == '9am' or '5pm' or '1am':
		upsStatusRow = upsStatusRow + 1
	elif hour == '10am' or '6pm' or '2am':
		upsStatusRow = upsStatusRow + 2
	elif hour == '11am' or '7pm' or '3am':
		upsStatusRow = upsStatusRow + 3
	elif hour == '12pm' or '8pm' or '4am':
		upsStatusRow = upsStatusRow + 4
	elif hour == '1pm' or '9pm' or '5am':
		upsStatusRow = upsStatusRow + 5
	elif hour == '2pm' or '10pm' or '6am':
		upsStatusRow = upsStatusRow + 6
	elif hour == '3pm' or '11pm' or '7am':
		upsStatusRow = upsStatusRow + 7
	else: 
		print 'error'	
		
	for i in activeUps: 																							#This loop is for getting temp and humitty from the user and writing them to spreadsheet.
		print "__________________________________________________________"
		
		print "Current UPS %s " % i

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
	
	ws.write( notesDesignationRow, notesDesignationColumn, "%s Out of tolerance units: %s" % (hour , notes))
	notesDesignationRow = notesDesignationRow + 1
	
	#saveAs = raw_input("Save file as: ") 																		#Save mechanism
	#saveAs = saveAs + ".xls"
	print "Saving as: %s" % saveAs
	
	while True:
		try:
			wb.save(saveAs)
			break
		except IOError:
			print "Cannot write to open sheet!"
			raw_input('Close the sheet then press any key to continue.')
			
	#print "Opening %s" % saveAs																					#Open the newly created spreadsheet
	#os.system("start %s" % saveAs)
	#print "Success!"
	
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
			