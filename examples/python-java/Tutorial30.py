"""----------------------------------------------------------
Tutorial 30

This tutorial shows how to export data to CSV file in Python.
----------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')
java_import(gateway.jvm,'EasyXLS.Constants.*')

print("Tutorial 30\n----------\n")

# Create an instance of the class that exports Excel files, having a sheet
workbook = gateway.jvm.ExcelDocument(2)

# Set the sheet name
workbook.easy_getSheetAt(0).setSheetName("First tab")

# Get the table of data for the worksheet
xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()

# Add data in cells for report header
for column in range(5):
	xlsFirstTable.easy_getCell(0,column).setValue("Column " + str(column + 1))
	xlsFirstTable.easy_getCell(0,column).setDataType(gateway.jvm.DataType.STRING)

# Add data in cells for report values
for row in range(100):
    for column in range(5):
        xlsFirstTable.easy_getCell(row+1,column).setValue("Data " + str(row + 1) + ", " + str(column + 1))
        xlsFirstTable.easy_getCell(row+1,column).setDataType(gateway.jvm.DataType.STRING)

# Export CSV file
print("Writing file C:\\Samples\\Tutorial30 - export CSV file..csv.")
workbook.easy_WriteCSVFile("C:\\Samples\\Tutorial30 - export CSV file..csv", "First tab")

# Confirm export of CSV file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()