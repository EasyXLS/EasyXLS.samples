<%@ Language=VBScript %>

<!-- #INCLUDE FILE="Alignment.inc" -->
<!-- #INCLUDE FILE="Border.inc" -->
<!-- #INCLUDE FILE="DataType.inc" -->
<!-- #INCLUDE FILE="Color.inc" -->
<%
	'========================================================================
	'Tutorial 06
	'
	' This code sample shows how to create an Excel file in Classic ASP with
	' multiple sheets. The first sheet is protected and
	' filled with data. The cells are formatted and locked.
	'========================================================================
	
	response.write("Tutorial 06<br>")
	response.write("----------<br>")

	' Create an instance of the class that creates Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create two worksheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
	
	' Protect first sheet
	workbook.easy_getSheetAt(0).setSheetProtected(true)
	
	' Get the table of data for the first worksheet
	Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()
	
	' Create the formatting style for the header
	set xlsStyleHeader = Server.CreateObject("EasyXLS.ExcelStyle")
	xlsStyleHeader.setFont("Verdana")
	xlsStyleHeader.setFontSize(8)
	xlsStyleHeader.setItalic(True)
	xlsStyleHeader.setBold(True)
	xlsStyleHeader.setForeground(CLng(COLOR_YELLOW))
	xlsStyleHeader.setBackground(CLng(COLOR_BLACK))
	xlsStyleHeader.setBorderColors CLng(COLOR_GRAY), CLng(COLOR_GRAY), CLng(COLOR_GRAY), CLng(COLOR_GRAY)
	xlsStyleHeader.setBorderStyles BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM
	xlsStyleHeader.setHorizontalAlignment(ALIGNMENT_ALIGNMENT_CENTER)
	xlsStyleHeader.setVerticalAlignment(ALIGNMENT_ALIGNMENT_BOTTOM)
	xlsStyleHeader.setWrap(True)
	xlsStyleHeader.setDataType(DATATYPE_STRING)
	
	' Add data in cells for report header
	for column = 0 to 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (column + 1))
		xlsFirstTable.easy_getCell(0,column).setStyle(xlsStyleHeader)
	next
	xlsFirstTable.easy_getRowAt(0).setHeight(30)
	
	' Create a formatting style for cells
	Set xlsStyleData = Server.CreateObject("EasyXLS.ExcelStyle")
	xlsStyleData.setHorizontalAlignment(ALIGNMENT_ALIGNMENT_LEFT)
	xlsStyleData.setForeground(CLng(COLOR_DARKGRAY))
	xlsStyleData.setWrap(False)
	' Protect cells
	xlsStyleData.setLocked(True)
	xlsStyleData.setDataType(DATATYPE_STRING)
	
	' Add data in cells for report values
	for row = 0 to 99
		for column = 0 to 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setStyle(xlsStyleData)
		next
	next

	' Set column widths
	xlsFirstTable.setColumnWidth_2 0, 70
	xlsFirstTable.setColumnWidth_2 1, 100
	xlsFirstTable.setColumnWidth_2 2, 70
	xlsFirstTable.setColumnWidth_2 3, 100
	xlsFirstTable.setColumnWidth_2 4, 70

	' Create Excel file
	response.write("Writing file: C:\Samples\Tutorial06 - protect Excel sheet.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial06 - protect Excel sheet.xlsx")
	
	' Confirm the creation of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
