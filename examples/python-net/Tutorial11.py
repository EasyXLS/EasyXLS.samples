"""-----------------------------------------------------------
Tutorial 11

This tutorial shows how to create an Excel file in Python that
has a cell that contains SUM formula for a range of cells.
-----------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *

print("Tutorial 11\n----------\n")

# Create an instance of the class that exports Excel files
workbook = ExcelDocument()

# Create a sheet
workbook.easy_addWorksheet("Formula")

# Get the table of data for the sheet, add data in sheet and the formula
xlsTable = workbook.easy_getSheet("Formula").easy_getExcelTable()
xlsTable.easy_getCell("A1").setValue("1")
xlsTable.easy_getCell("A2").setValue("2")
xlsTable.easy_getCell("A3").setValue("3")
xlsTable.easy_getCell("A4").setValue("4")
xlsTable.easy_getCell("A6").setValue("=SUM(A1:A4)")

# Export Excel file
print("Writing file C:\\Samples\\Tutorial11 - formulas in Excel.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial11 - formulas in Excel.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()