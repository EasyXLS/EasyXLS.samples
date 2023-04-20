"""---------------------------------------------------------------
Tutorial 03
This tutorial shows how to create an Excel file that has
multiple sheets in Python. The created Excel file is empty and the
next tutorial shows how to add data into sheets. 
---------------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *

print("Tutorial 03\n-----------\n")

# Create an instance of the class that creates Excel files, having two sheets
workbook = ExcelDocument(2)

# Set the sheet names
workbook.easy_getSheetAt(0).setSheetName("First tab")
workbook.easy_getSheetAt(1).setSheetName("Second tab")

# Create the Excel file
print("Writing file C:\\Samples\\Tutorial03 - create Excel file.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial03 - create Excel file.xlsx")

# Confirm the creation of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()