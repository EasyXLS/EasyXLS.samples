"""--------------------------------------------------------------------------------
Tutorial 15

This tutorial shows how to create an Excel file with hyperlinks in Python.

EasyXLS supports the following hyperlink types:
    (1) - hyperlink to URL
    (2) - hyperlink to file
    (3) - hyperlink to UNC
    (4) - hyperlink to cell in the same Excel file
    (5) - hyperlink to name 

Every type of hyperlink accepts a tool tip description.

Every type of hyperlink accepts a text mark. A text mark is a link inside the file.
--------------------------------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')
java_import(gateway.jvm,'EasyXLS.Constants.*')

print("Tutorial 15\n-----------\n")

# Create an instance of the class that exports Excel files having two sheets
workbook = gateway.jvm.ExcelDocument(2)

# Set the sheet names
xlsTab1 = workbook.easy_getSheetAt(0)
xlsTab2 = workbook.easy_getSheetAt(1)
xlsTab1.setSheetName("First tab")
xlsTab2.setSheetName("Second tab")

# Create hyperlink to URL
xlsTab1.easy_addHyperlink(gateway.jvm.HyperlinkType.URL, "http://www.euoutsourcing.com", 
                          "Link to URL", "B2:E2")

# Create hyperlink to file
xlsTab1.easy_addHyperlink(gateway.jvm.HyperlinkType.FILE, "c:\\myfile.xls", "Link to file", "B3")

# Create hyperlink to UNC
xlsTab1.easy_addHyperlink(gateway.jvm.HyperlinkType.UNC, "\\\\computerName\\Folder\\file.txt", 
                          "Link to UNC", "B4:D4")

# Create hyperlink to cell on second sheet
xlsTab1.easy_addHyperlink(gateway.jvm.HyperlinkType.CELL, "'Second tab'!D3", "Link to CELL", "B5")

# Create a name on the second sheet
xlsTab2.easy_addName("Name", "=Second tab!$A$1:$A$4")

# Create hyperlink to name
xlsTab1.easy_addHyperlink(gateway.jvm.HyperlinkType.CELL, "Name", "Link to a name", "B6")

# Export Excel file
print("Writing file C:\\Samples\\Tutorial15 - hyperlinks in Excel.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial15 - hyperlinks in Excel.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()