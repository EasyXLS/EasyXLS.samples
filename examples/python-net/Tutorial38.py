"""----------------------------------------------------------------------------------
Tutorial 38

This tutorial shows how to read an Excel XLSB file in Python
(the XLSB file generated by Tutorial 29 as base template),
modify some data and save it to another XLSB file (Tutorial38 - read XLSB file.xlsb).
----------------------------------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *

print("Tutorial 38\n-----------\n")

# Create an instance of the class that reads Excel files
workbook = ExcelDocument()

# Read XLSB file
print("Reading file C:\\Samples\\Tutorial29.xlsb")
		
if workbook.easy_LoadXLSBFile("C:\\Samples\\Tutorial29.xlsb"):
    # Get the table of data for the second worksheet
	xlsSecondTable = workbook.easy_getSheet("Second tab").easy_getExcelTable()

    # Write some data to the second sheet
	xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial38")
			
	for column in range(5):
		xlsSecondTable.easy_getCell(1, column).setValue("Data " + str(column + 1))

    # Export the new XLSB file
	print("\nWriting file C:\\Samples\\Tutorial38 - read XLSB file.xlsb.")
	workbook.easy_WriteXLSBFile("C:\\Samples\\Tutorial38 - read XLSB file.xlsb")

    # Confirm export of Excel file
	sError = workbook.easy_getError()

	if sError == "":
		print("\nFile successfully created.\n\n")
	else:
            print("\nError encountered: " + sError + "\n\n")
else:
	print("\nError reading file C:\\Samples\\Tutorial29.xlsb" + workbook.easy_getError() + "\n\n")	
		
# Dispose memory
gc.collect()