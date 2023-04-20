"""-------------------------------------------------------------------------
Tutorial 33

This tutorial shows how to set document properties for Excel file in Python, 
like 'Subject' property for summary information, 'Manager' property for
document summary information and a custom property.
-------------------------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *
from EasyXLS.Constants import *

print("Tutorial 33\n-----------\n")

# Create an instance of the class that exports Excel files
workbook = ExcelDocument(1)

# Set the 'Subject' document property
workbook.getSummaryInformation().setSubject("This is the subject")

# Set the 'Manager' document property
workbook.getDocumentSummaryInformation().setManager("This is the manager")

# Set a custom document property
workbook.getDocumentSummaryInformation().setCustomProperty("PropertyName", FileProperty.VT_NUMBER, "4")

# Export Excel file
print("Writing file C:\\Samples\\Tutorial33 - Excel file properties.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial33 - Excel file properties.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()