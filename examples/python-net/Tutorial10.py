"""--------------------------------------------------------------------------------
Tutorial 10
 
This tutorial shows how to export an Excel file with a merged cell range in Python.
--------------------------------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *

print("Tutorial 10\n-----------\n")

# Create an instance of the class that exports Excel files
workbook = ExcelDocument(1)

# Get the table of data for the worksheet
xlsTable = workbook.easy_getSheet("Sheet1").easy_getExcelTable()

# Merge cells by range
xlsTable.easy_mergeCells("A1:C3")

# Export Excel file
print("Writing file C:\\Samples\\Tutorial10 - merge cells in Excel.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial10 - merge cells in Excel.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()	