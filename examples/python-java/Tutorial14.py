"""-------------------------------------------------------
Tutorial 14

This tutorial shows how to create an Excel file in Python
having a sheet and conditional formatting for cell ranges.
-------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')
java_import(gateway.jvm,'EasyXLS.Constants.*')
java_import(gateway.jvm,'java.awt.Color')

print("Tutorial 14\n-----------\n")

# Create an instance of the class that exports Excel files having one sheet
workbook = gateway.jvm.ExcelDocument(1)

# Get the table of data for the first worksheet
xlsTab = workbook.easy_getSheet("Sheet1")
xlsTable = xlsTab.easy_getExcelTable()

# Add data in cells
for i in range(6):
    for j in range(4):
        if i<2 and j<2:
            xlsTable.easy_getCell(i, j).setValue("12")
        elif j==2 and i<2:
            xlsTable.easy_getCell(i, j).setValue("1000")
        else:
            xlsTable.easy_getCell(i, j).setValue("9")
        xlsTable.easy_getCell(i, j).setDataType(gateway.jvm.DataType.NUMERIC)

# Set conditional formatting
xlsTab.easy_addConditionalFormatting("A1:C3", gateway.jvm.ConditionalFormatting.OPERATOR_BETWEEN, 
                                     "=9", "=11", True, True, gateway.jvm.Color.RED)

# Set another conditional formatting
xlsTab.easy_addConditionalFormatting("A6:C6", gateway.jvm.ConditionalFormatting.OPERATOR_BETWEEN, 
                                     "=COS(PI())+2", "", gateway.jvm.Color.ORANGE)
xlsTab.easy_getConditionalFormattingAt("A6:C6").getConditionAt(0).setConditionType_
(gateway.jvm.ConditionalFormatting.CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA)

# Export Excel file
print("Writing file C:\\Samples\\Tutorial14 - conditional formatting in Excel.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial14 - conditional formatting in Excel.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()