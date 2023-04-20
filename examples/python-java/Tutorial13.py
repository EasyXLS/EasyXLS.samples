"""-------------------------------------------------------------
Tutorial 13

This tutorial shows how to create an Excel file in Python having
multiple sheets. The second sheet contains a named area range.
The A1:A10 cell range contains data validators, drop down list 
and whole number validation.
-------------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')
java_import(gateway.jvm,'EasyXLS.Constants.*')

print("Tutorial 13\n----------\n")

# Create an instance of the class that exports Excel files, having two sheets
workbook = gateway.jvm.ExcelDocument(2)

# Set the sheet names
workbook.easy_getSheetAt(0).setSheetName("First tab")
workbook.easy_getSheetAt(1).setSheetName("Second tab")

# Get the table of data for the second worksheet and populate the worksheet
xlsSecondTab = workbook.easy_getSheetAt(1)
xlsSecondTable = xlsSecondTab.easy_getExcelTable()
xlsSecondTable.easy_getCell("A1").setValue("Range data 1")
xlsSecondTable.easy_getCell("A2").setValue("Range data 2")
xlsSecondTable.easy_getCell("A3").setValue("Range data 3")
xlsSecondTable.easy_getCell("A4").setValue("Range data 4")

# Create a named area range
xlsSecondTab.easy_addName("Range", "=Second tab!$A$1:$A$4")

# Add data validation as drop down list type
xlsFirstTab = workbook.easy_getSheetAt(0)
xlsFirstTab.easy_addDataValidator("A1:A10", gateway.jvm.DataValidator.VALIDATE_LIST, 
                                  gateway.jvm.DataValidator.OPERATOR_EQUAL_TO, "=Range", "")

# Add data validation as whole number type
xlsFirstTab.easy_addDataValidator("B1:B10", gateway.jvm.DataValidator.VALIDATE_WHOLE_NUMBER, 
                                  gateway.jvm.DataValidator.OPERATOR_BETWEEN, "=4", "=100")

# Export Excel file
print("Writing file C:\\Samples\\Tutorial13 - cell validation in Excel.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial13 - cell validation in Excel.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()