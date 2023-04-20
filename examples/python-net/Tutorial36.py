"""----------------------------------------------------------------------------------
Tutorial 36

This tutorial shows how to read an Excel XLSX file in Python
(the XLSX file generated by Tutorial 04 as base template),
modify some data and save it to another XLSX file (Tutorial36 - read XLSX file.xlsx).
----------------------------------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *

print("Tutorial 36\n-----------\n")

# Create an instance of the class that reads Excel files
workbook = ExcelDocument()

# Read XLSX file
print("Reading file C:\\Samples\\Tutorial04.xlsx")
		
if workbook.easy_LoadXLSXFile("C:\\Samples\\Tutorial04.xlsx"):
    # Get the table of data for the second worksheet
	xlsSecondTable = workbook.easy_getSheet("Second tab").easy_getExcelTable()

    # Write some data to the second sheet
	xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial36")
			
	for column in range(5):
		xlsSecondTable.easy_getCell(1, column).setValue("Data " + str(column + 1))

    # Export the new XLSX file
	print("\nWriting file C:\\Samples\\Tutorial36 - read XLSX file.xlsx.")
	workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial36 - read XLSX file.xlsx")

    # Confirm export of Excel file
	sError = workbook.easy_getError()

	if sError == "":
		print("\nFile successfully created.\n\n")
	else:
            print("\nError encountered: " + sError + "\n\n")
else:
	print("\nError reading file C:\\Samples\\Tutorial04.xlsx" + workbook.easy_getError() + "\n\n")

# Dispose memory
gc.collect()