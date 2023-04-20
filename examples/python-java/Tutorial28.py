"""------------------------------------------------------------
Tutorial 28
This tutorial shows how to export data to XLS file that has
multiple sheets in Python. The first sheet is filled with data.
------------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')
java_import(gateway.jvm,'EasyXLS.Constants.*') 

print("Tutorial 28\n----------\n")

# Create an instance of the class that exports Excel files, having two sheets
workbook = gateway.jvm.ExcelDocument(2)

# Set the sheet names
workbook.easy_getSheetAt(0).setSheetName("First tab")
workbook.easy_getSheetAt(1).setSheetName("Second tab")

# Get the table of data for the first worksheet
xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()

# Add data in cells for report header
for column in range(5):
    xlsFirstTable.easy_getCell(0,column).setValue("Column " + str(column + 1))
    xlsFirstTable.easy_getCell(0,column).setDataType(gateway.jvm.DataType.STRING)
xlsFirstTable.easy_getRowAt(0).setHeight(30)

# Add data in cells for report values
for row in range(100):
    for column in range(5):
        xlsFirstTable.easy_getCell(row+1,column).setValue("Data " + str(row + 1) + ", " + str(column + 1))
        xlsFirstTable.easy_getCell(row+1,column).setDataType(gateway.jvm.DataType.STRING)

# Set column widths
xlsFirstTable.setColumnWidth(0, 70)
xlsFirstTable.setColumnWidth(1, 100)
xlsFirstTable.setColumnWidth(2, 70)
xlsFirstTable.setColumnWidth(3, 100)
xlsFirstTable.setColumnWidth(4, 70)

# Export the XLS file
print("Writing file C:\\Samples\\Tutorial28 - export XLS file.xls.")
workbook.easy_WriteXLSFile("C:\\Samples\\Tutorial28 - export XLS file.xls")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()