"""------------------------------------------------------
Tutorial 20

This tutorial shows how to create an Excel file in Python 
and apply an auto-filter to a range of cells.
------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *
from EasyXLS.Constants import *

print("Tutorial 20\n-----------\n")

# Create an instance of the class that exports Excel files having one sheet
workbook = ExcelDocument(1)

# Get the table of data for the worksheet 
xlsTab = workbook.easy_getSheet("Sheet1")
xlsTable = xlsTab.easy_getExcelTable()

# Add data in cells for report header
for column in range(5):
	xlsTable.easy_getCell(0,column).setValue("Column " + str(column + 1))
	xlsTable.easy_getCell(0,column).setDataType(DataType.STRING)

# Add data in cells for report values
for row in range(100):
    for column in range(5):
        xlsTable.easy_getCell(row+1,column).setValue("Data " + str(row + 1) + ", " + str(column + 1))
        xlsTable.easy_getCell(row+1,column).setDataType(DataType.STRING)

# Apply auto-filter on cell range A1:E1
xlsFilter = xlsTab.easy_getFilter()
xlsFilter.setAutoFilter("A1:E1")

# Export Excel file
print("Writing file C:\\Samples\\Tutorial20 - autofilter in Excel sheet.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial20 - autofilter in Excel sheet.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()