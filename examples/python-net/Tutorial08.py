"""------------------------------------------------------------------------------
Tutorial 08

This tutorial shows how to create an Excel file in Python having multiple sheets. 
The first sheet is filled with data and the cells are formatted and locked.
The column header has comments.
The first sheet has header & footer.
------------------------------------------------------------------------------"""

import clr
import gc

clr.AddReference('EasyXLS')
from EasyXLS import *
from System.Drawing import *
from EasyXLS.Constants import *

print("Tutorial 08\n-----------\n")

# Create an instance of the class that exports Excel files, having two sheets
workbook = ExcelDocument(2)

# Set the sheet names
workbook.easy_getSheetAt(0).setSheetName("First tab")
workbook.easy_getSheetAt(1).setSheetName("Second tab")

# Protect first sheet
workbook.easy_getSheetAt(0).setSheetProtected(True)

# Get the table of data for the first worksheet
xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()

# Create the formatting style for the header
xlsStyleHeader = ExcelStyle("Verdana", 8, True, True, Color.Yellow)		
xlsStyleHeader.setBackground(Color.Black)
xlsStyleHeader.setBorderColors(Color.Gray, Color.Gray, Color.Gray, Color.Gray)
xlsStyleHeader.setBorderStyles(Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM)
xlsStyleHeader.setHorizontalAlignment(Alignment.ALIGNMENT_CENTER)
xlsStyleHeader.setVerticalAlignment(Alignment.ALIGNMENT_BOTTOM)
xlsStyleHeader.setWrap(True)
xlsStyleHeader.setDataType(DataType.STRING)

# Add data in cells for report header
for column in range(5):
	xlsFirstTable.easy_getCell(0, column).setValue("Column " + str(column + 1))
	xlsFirstTable.easy_getCell(0, column).setStyle(xlsStyleHeader)

    # Add comment for report header cells
	xlsFirstTable.easy_getCell(0, column).setComment("This is column no " + str(column + 1))

xlsFirstTable.easy_getRowAt(0).setHeight(30)

# Add data in cells for report values
for row in range(100):
    for column in range(5):
        xlsFirstTable.easy_getCell(row+1, column).setValue("Data " + str(row + 1) + ", " + str(column + 1))

# Create a formatting style for cells
xlsStyleData = ExcelStyle()
xlsStyleData.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)
xlsStyleData.setForeground(Color.DarkGray)
xlsStyleData.setWrap(False)
xlsStyleData.setDataType(DataType.STRING)
xlsStyleData.setLocked(True)
xlsFirstTable.easy_setRangeStyle("A2:E101", xlsStyleData)

# Set column widths
xlsFirstTable.setColumnWidth(0, 70)
xlsFirstTable.setColumnWidth(1, 100)
xlsFirstTable.setColumnWidth(2, 70)
xlsFirstTable.setColumnWidth(3, 100)
xlsFirstTable.setColumnWidth(4, 70)

# Add header on center section
xlsFirstTab = workbook.easy_getSheetAt(0)
xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertSingleUnderline()
xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertFile()
xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertValue(" - How to create header and footer")

# Add header on right section
xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertDate()
xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertValue(" ")
xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertTime()

# Add footer on center section
xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertPage()
xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertValue(" of ")
xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertPages()

# Export Excel file
print("Writing file C:\\Samples\\Tutorial08 - header and footer in Excel.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial08 - header and footer in Excel.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()