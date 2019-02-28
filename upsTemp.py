import xlwt
from datetime import datetime
import os


			
wb = xlwt.Workbook()
ws = wb.add_sheet("A test sheet")
ws.write( 0, 1, "test")


activeUps = ['UPS 130-1', 'UPS 130-2', 'UPS 135-1', 'UPS 135-2', 'UPS 135-3', 'UPS 240-1', 'UPS 240-2']

upsStatusDesignationRow = 32
upsStatusDesignationColumn = 9
upsIdDesignationRow = 31
upsIdDesignationColumn = 9
upsStatusRow = 33
upsStatusColumn = 9

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
		currentUpsStatus = 'System Operational'
	elif onOff == 'n':
		currentUpsStatus = 'System Non-operational'
	else:
		print "Error"
		currentUpsStatus = "Error"

	ws.write( upsStatusRow, upsStatusColumn, currentUpsStatus)
	upsStatusColumn = upsStatusColumn + 2
	
print "__________________________________________________________"
print "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
	
saveAs = raw_input("Save file as: ") 																		#Save mechanism
saveAs = saveAs + ".xls"
print "Saving as: %s" % saveAs
wb.save(saveAs) 																							#May want to automate this at some point for easy of use and standardization reasons.

print "Opening %s" % saveAs																					#Open the newly created spreadsheet
os.system("start %s" % saveAs)
print "Success!"
