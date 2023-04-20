"""----------------------------------------------------------
Tutorial 34

This tutorial shows how to import Excel to DataSet in Python. 
The data is imported from the active sheet of the Excel file
(the Excel file generated in Tutorial 09).
----------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *

print("Tutorial 34\n-----------\n")

# Create an instance of the class that imports Excel files
workbook = ExcelDocument()

# Import Excel file to DataSet
print("Reading file C:\\Samples\\Tutorial09.xlsx.\n")
ds = workbook.easy_ReadXLSXActiveSheet_AsDataSet("C:\\Samples\\Tutorial09.xlsx")

# Display imported DataSet values
dt = ds.Tables[0]
for row in range(dt.Rows.Count):
    for column in range(dt.Columns.Count):
        print("At row " + str(row + 1) + ", column " + str(column + 1) +
            " the value is '" + dt.Rows[row].ItemArray[column] + "'")

# Dispose memory
gc.collect()