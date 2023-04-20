"""---------------------------------------------------------------------------
Tutorial 17

This tutorial shows how to create an Excel file with groups on rows in Python.
The Excel file has two worksheets. The first one is full with data and
contains the data groups.
---------------------------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *
from EasyXLS.Constants import *

print("Tutorial 17\n----------\n")

# Create an instance of the class that exports Excel files having two sheets
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

xlsFirstTable.easy_getRowAt(0).setHeight(30)

# Add data in cells for report values
for row in range(25):
    for column in range(5):
        xlsFirstTable.easy_getCell(row+1,column).setValue("Data " + str(row + 1) + ", " + str(column + 1))
        xlsFirstTable.easy_getCell(row+1,column).setDataType(DataType.STRING)

# Set column widths
xlsFirstTable.setColumnWidth(0, 70)
xlsFirstTable.setColumnWidth(1, 100)
xlsFirstTable.setColumnWidth(2, 70)
xlsFirstTable.setColumnWidth(3, 100)
xlsFirstTable.setColumnWidth(4, 70)

# Group rows and format A1:E26 cell range
xlsFirstDataGroup = ExcelDataGroup("A1:E26", DataGroup.GROUP_BY_ROWS, False)
xlsFirstDataGroup.setAutoFormat(ExcelAutoFormat(Styles.AUTOFORMAT_EASYXLS1))
workbook.easy_getSheetAt(0).easy_addDataGroup(xlsFirstDataGroup)

# Group rows and format A2:E10 cell range, outline level two, inside previous group
xlsSecondDataGroup = ExcelDataGroup("A2:E10", DataGroup.GROUP_BY_ROWS, False)		
xlsSecondDataGroup.setAutoFormat(ExcelAutoFormat(Styles.AUTOFORMAT_EASYXLS2))		 
workbook.easy_getSheetAt(0).easy_addDataGroup(xlsSecondDataGroup)

# Export Excel file
print("Writing file C:\\Samples\\Tutorial17 - group data in Excel.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial17 - group data in Excel.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()