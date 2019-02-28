import xlwt
from datetime import datetime
import os
import time 



wb = xlwt.Workbook()
ws = wb.add_sheet("A test sheet") 																			#this is to create the sheet but I need to change this to something relevant later



unitIdRow = 1 																									#This is the first Row/Column to be written to for the crac unit's ID.
unitIdColumn = 1
tempDesignationRow = 2																							#This is the first Row/Column to be written to for the Designation of temperature column's.
tempDesignationColumn = 1
humittyDesignationRow = 2																						#This is the first Row/Column to be written to for the Designation of humitty column's.
humittyDesignationColumn = 2
timeDesignationRow = 3																							#This is the first Row/Column to be written to for the Designation of the time of day.
timeDesignationColumn = 0

activeUnits = ['CRAC 135-1', 'CRAC 135-2', 'CRAC 135-3', 'CRAC 135-4', 'CRAC 135-5', 'CRAC 135-6', 'CRAC 135-7', 'CRAC 135-8', 'CRAC 135-9', 'CRAC 135-10', 'CRAC 135-11']
timeFirstShift = ['8:00 AM', '9:00 AM', '10:00 AM', '11:00 AM', '12:00 PM', '1:00 PM', '2:00 PM', '3:00 PM']

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

	
saveAs = raw_input("Save file as: ") 																		#Save mechanism
saveAs = saveAs + ".xls"
print "Saving as: %s" % saveAs
wb.save(saveAs)
print "Opening %s" % saveAs																					#Open the newly created spreadsheet
os.system("start %s" % saveAs)
print "Success!"