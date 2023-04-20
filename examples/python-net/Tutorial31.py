"""-----------------------------------------------------------
Tutorial 31

This tutorial shows how to export data to HTML file in Python.
-----------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *
from EasyXLS.Constants import *

print("Tutorial 31\n----------\n")

# Create an instance of the class that exports Excel files, having a sheet
workbook = ExcelDocument(1)

# Set the sheet name
workbook.easy_getSheetAt(0).setSheetName("First tab")

# Get the table of data for the worksheet
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

# Export HTML file
print("Writing file C:\\Samples\\Tutorial31 - export HTML file.html.")
workbook.easy_WriteHTMLFile( "C:\\Samples\\Tutorial31 - export HTML file.html", "First tab")

# Confirm export of HTML file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()