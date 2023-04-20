"""------------------------------------------------------------
Tutorial 34

This tutorial shows how to import Excel to ResultSet in Python.
The data is imported from the active sheet of the Excel file
(the Excel file generated in Tutorial 09).
------------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')
java_import(gateway.jvm,'java.io.FileInputStream')

print("Tutorial 34\n-----------\n")

# Create an instance of the class that imports Excel files
workbook = gateway.jvm.ExcelDocument()

# Import Excel file to ResultSet
print("Reading file C:\\Samples\\Tutorial09.xlsx.\n")

file = gateway.jvm.FileInputStream("C:\\Samples\\Tutorial09.xlsx")
rs = workbook.easy_ReadXLSXActiveSheet_AsResultSet(file)

# Display imported ResultSet values
columnCount = rs.getMetaData().getColumnCount()

row = 0
while rs.next():
    for column in range(columnCount):
        print("At row " + str(row + 1) + ", column " + str(column+1) +
              " the value is '" + rs.getString(column+1) + "'")
    row=row+1

# Dispose memory
gc.collect()