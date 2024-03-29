"""-------------------------------------------------------------------------------
Tutorial 37

This tutorial shows how to read an Excel XLS file in Python
(the XLS file generated by Tutorial 28 as base template),
modify some data and save it to another XLS file (Tutorial37 - read XLS file.xls).
-------------------------------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')

print("Tutorial 37\n-----------\n")

# Create an instance of the class that reads Excel files
workbook = gateway.jvm.ExcelDocument()

# Read XLS file
print("Reading file C:\\Samples\\Tutorial28.xls")
		
if workbook.easy_LoadXLSFile("C:\\Samples\\Tutorial28.xls"):
    # Get the table of data for the second worksheet
	xlsSecondTable = workbook.easy_getSheet("Second tab").easy_getExcelTable()

    # Write some data to the second sheet
	xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial37")
			
	for column in range(5):
		xlsSecondTable.easy_getCell(1, column).setValue("Data " + str(column + 1))

    # Export the new XLS file
	print("\nWriting file C:\\Samples\\Tutorial37 - read XLS file.xls.")
	workbook.easy_WriteXLSFile("C:\\Samples\\Tutorial37 - read XLS file.xls")

    # Confirm export of Excel file
	sError = workbook.easy_getError()

	if sError == "":
		print("\nFile successfully created.\n")
	else:
            print("\nError encountered: " + sError + "\n")
else:
	print("Error reading file C:\\Samples\\Tutorial28.xls")

# Dispose memory
gc.collect()