"""---------------------------------------------------------------
Tutorial 02

This code sample shows how to export List to Excel file in Python.
The List contains data from a SQL database
The cells are formatted using a user-defined format.
---------------------------------------------------------------"""

import gc
import sqlite3

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')
java_import(gateway.jvm,'EasyXLS.Constants.*')
java_import(gateway.jvm,'EasyXLS.Util.*')
java_import(gateway.jvm,'java.awt.Color')

print("Tutorial 02\n-----------\n")

# Create an instance of the class that exports Excel files
workbook = gateway.jvm.ExcelDocument()

# Create the database connection
conn = sqlite3.connect('chinook.db')

# Query the database
cursor = conn.execute("SELECT strftime('%m/%d/%Y', I.InvoiceDate), C.FirstName || ' ' || C.LastName, I.BillingAddress, I.BillingCity, I.BillingState, I.BillingCountry, I.Total FROM Invoices I INNER JOIN Customers C ON I.CustomerId=C.CustomerId LIMIT 100")
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

# Create an instance of the class used to format the cells in the report
xlsAutoFormat = gateway.jvm.ExcelAutoFormat()
        
# Set the formatting style of the header
xlsHeaderStyle = gateway.jvm.ExcelStyle(gateway.jvm.Color(144, 238, 144))
xlsHeaderStyle.setFontSize(12)
xlsAutoFormat.setHeaderRowStyle(xlsHeaderStyle)

# Set the formatting style of the cells (alternating style)
xlsEvenRowStripesStyle = gateway.jvm.ExcelStyle(gateway.jvm.Color(255, 250, 240))
xlsEvenRowStripesStyle.setFormat("$0.00")
xlsEvenRowStripesStyle.setHorizontalAlignment(gateway.jvm.Alignment.ALIGNMENT_LEFT)
xlsAutoFormat.setEvenRowStripesStyle(xlsEvenRowStripesStyle)
xlsOddRowStripesStyle = gateway.jvm.ExcelStyle(gateway.jvm.Color(240, 247, 239))
xlsOddRowStripesStyle.setFormat("$0.00")
xlsOddRowStripesStyle.setHorizontalAlignment (gateway.jvm.Alignment.ALIGNMENT_LEFT)
xlsAutoFormat.setOddRowStripesStyle(xlsOddRowStripesStyle)
xlsLeftColumnStyle = gateway.jvm.ExcelStyle(gateway.jvm.Color(255, 250, 240))
xlsLeftColumnStyle.setFormat("mm/dd/yyyy")
xlsLeftColumnStyle.setHorizontalAlignment (gateway.jvm.Alignment.ALIGNMENT_LEFT)
xlsAutoFormat.setLeftColumnStyle(xlsLeftColumnStyle)

# Export the Excel file
print("Writing file C:\\Samples\\Tutorial02 - export List to Excel with formatting.xlsx.")
workbook.easy_WriteXLSXFile_FromList("c:\\Samples\\Tutorial02 - export List to Excel with formatting.xlsx", lstRows, xlsAutoFormat, "Sheet1")

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