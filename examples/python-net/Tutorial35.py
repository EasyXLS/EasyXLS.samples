"""----------------------------------------------------------------
Tutorial 35

This tutorial shows how to import Excel sheet to DataSet in Python.
The data is imported from a specific Excel sheet (For this example 
we use the Excel file generated in Tutorial 09).
----------------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *

print("Tutorial 35\n-----------\n")

# Create an instance of the class that imports Excel files
workbook = ExcelDocument()

# Import Excel sheet to DataSet
print("Reading file C:\\Samples\\Tutorial09.xlsx.\n")
ds = workbook.easy_ReadXLSXSheet_AsDataSet("C:\\Samples\\Tutorial09.xlsx", "First tab")

# Display imported DataSet values
dt = ds.Tables[0]
for row in range(dt.Rows.Count):
    for column in range(dt.Columns.Count):
        print("At row " + str(row + 1) + ", column " + str(column + 1) +
        " the value is '" + dt.Rows[row].ItemArray[column] + "'")
 
# Dispose memory
gc.collect()