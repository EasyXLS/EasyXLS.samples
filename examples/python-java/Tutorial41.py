"""------------------------------------------------------------------------
Tutorial 41

This tutorial shows how to convert XML spreadsheet to Excel in Python. The
XML Spreadsheet generated by Tutorial 32 is imported, some data is modified
and after that is exported as Excel file.
------------------------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')

print("Tutorial 41\n-----------\n")

# Create an instance of the class used to import/export Excel files
workbook = gateway.jvm.ExcelDocument()

# Import XML Spreadsheet file
print("Reading file C:\\Samples\\Tutorial32.xml")
		
if workbook.easy_LoadXMLSpreadsheetFile("C:\\Samples\\Tutorial32.xml"):
    # Get the table of data from the second sheet and add some data in cells (optional step)
	xlsSecondTable = workbook.easy_getSheet("Second tab").easy_getExcelTable()
	xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial41")
						
	for column in range(5):
		xlsSecondTable.easy_getCell(1, column).setValue("Data " + str(column + 1))

    # Export Excel file
	print("\nWriting file C:\\Samples\\Tutorial41 - convert XML spreadsheet to Excel.xlsx.")
	workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial41 - convert XML spreadsheet to Excel.xlsx")

    # Confirm conversion of XML Spreadsheet to Excel
	sError = workbook.easy_getError()

	if sError == "":
		print("\nFile successfully created.\n")
	else:
            print("\nError encountered: " + sError + "\n")
else:
	print("\nError reading file C:\\Samples\\Tutorial32.xml \n" + workbook.easy_getError())
		
# Dispose memory
gc.collect()