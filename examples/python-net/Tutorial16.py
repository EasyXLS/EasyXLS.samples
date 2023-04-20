"""------------------------------------------------------------------
Tutorial 16

This tutorial shows how to create an Excel file with image in Python.
The Excel file has multiple sheets.
The first sheet has an image inserted.
------------------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *

print("Tutorial 16\n----------\n")

# Create an instance of the class that exports Excel files having two sheets
workbook = ExcelDocument(2)

# Set the sheet names
workbook.easy_getSheetAt(0).setSheetName("First tab")
workbook.easy_getSheetAt(1).setSheetName("Second tab")

# Insert image into sheet
workbook.easy_getSheetAt(0).easy_addImage("C:\\Samples\\EasyXLSLogo.JPG", "A1")

# Export Excel file
print("Writing file C:\\Samples\\Tutorial16 - images in Excel.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial16 - images in Excel.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()