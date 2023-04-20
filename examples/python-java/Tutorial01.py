"""--------------------------------------------------------------------
Tutorial 01

This code sample shows how to export List to Excel file in Python.
The List contains data from a SQL database
The cells are formatted using a predefined format.
--------------------------------------------------------------------"""

import gc
import sqlite3

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')
java_import(gateway.jvm,'EasyXLS.Constants.*')
java_import(gateway.jvm,'EasyXLS.Util.*') 

print("Tutorial 01\n-----------\n")

# Create an instance of the class that exports Excel files
workbook = gateway.jvm.ExcelDocument()

# Create the database connection
conn = sqlite3.connect('chinook.db')

# Query the database
cursor = conn.execute("SELECT strftime('%m/%d/%Y', I.InvoiceDate), C.FirstName || ' ' || C.LastName, "
                      "I.BillingAddress, I.BillingCity, I.BillingState, I.BillingCountry, "
                      "I.Total FROM Invoices I INNER JOIN Customers C ON "
                      "I.CustomerId=C.CustomerId LIMIT 100")
rows = cursor.fetchall()

# Create the list that stores the query values
lstRows = gateway.jvm.List()

# Add the report header row to the list
lstHeaderRow = gateway.jvm.List()
lstHeaderRow.addElement("Invoice date")
lstHeaderRow.addElement("Customer name")
lstHeaderRow.addElement("Billing address")
lstHeaderRow.addElement("Billing city")
lstHeaderRow.addElement("Billing state")
lstHeaderRow.addElement("Billing country")	
lstHeaderRow.addElement("Total")	
lstRows.addElement(lstHeaderRow)

# Add the query values from the database to the list
for row in rows:
    RowList = gateway.jvm.List()
    RowList.addElement(row[0])
    RowList.addElement(row[1])	
    RowList.addElement(row[2])
    RowList.addElement(row[3])
    RowList.addElement(row[4])
    RowList.addElement(row[5])
    RowList.addElement(row[6])
    lstRows.addElement(RowList)

# Export list to Excel file
print("Writing file C:\\Samples\\Tutorial01 - export List to Excel.xlsx.")
workbook.easy_WriteXLSXFile_FromList("c:\\Samples\\Tutorial01 - export List to Excel.xlsx", lstRows, gateway.jvm.ExcelAutoFormat(gateway.jvm.Styles.AUTOFORMAT_EASYXLS1), "Sheet1")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Close the database connection
cursor.close()
conn.close()

# Dispose memory
gc.collect()