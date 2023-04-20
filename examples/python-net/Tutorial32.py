"""----------------------------------------------------------------------
Tutorial 32

This tutorial shows how to export data to XML Spreadsheet file in Python.
----------------------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *
from EasyXLS.Constants import *

print("Tutorial 32\n----------\n")

# Create an instance of the class that exports Excel files, having two sheets
workbook = ExcelDocument(2)

# Set the sheet names
workbook.easy_getSheetAt(0).setSheetName("First tab")
workbook.easy_getSheetAt(1).setSheetName("Second tab")

# Get the table of data for the first worksheet
xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()

# Add data in cells for report header
for column in range(5):
	xlsFirstTable.easy_getCell(0,column).setValue("Column " + str(column + 1))
	xlsFirstTable.easy_getCell(0,column).setDataType(DataType.STRING)

# Add data in cells for report values
for row in range(100):
    for column in range(5):
        xlsFirstTable.easy_getCell(row+1,column).setValue("Data " + str(row + 1) + ", " + str(column + 1)) 
        xlsFirstTable.easy_getCell(row+1,column).setDataType(DataType.STRING)

# Apply a predefined format to the cells
xlsFirstTable.easy_setRangeAutoFormat("A1:E101", ExcelAutoFormat(Styles.AUTOFORMAT_EASYXLS1))

# Export XML Spreadsheet file
print("Writing file C:\\Samples\\Tutorial32 - export XML spreadsheet file.xml.")
workbook.easy_WriteXMLFile("C:\\Samples\\Tutorial32 - export XML spreadsheet file.xml")

# Confirm export of XML file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()