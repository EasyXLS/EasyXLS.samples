"""--------------------------------------------------------------------------
Tutorial 27

This tutorial shows how to create an Excel file in Python and
encrypt the Excel file by setting the password required for opening the file.
--------------------------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')

print("Tutorial 27\n----------\n")

# Create an instance of the class that exports Excel files, having two sheets
workbook = gateway.jvm.ExcelDocument(2)

# Set the sheet names
workbook.easy_getSheetAt(0).setSheetName("First tab")
workbook.easy_getSheetAt(1).setSheetName("Second tab")

# Set the password for protecting the Excel file when the file is open
workbook.easy_getOptions().setPasswordToOpen("password")

# Export Excel file
print("Writing file C:\\Samples\\Tutorial27 - protect Excel with password and encryption.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial27 - protect Excel with password and encryption.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()