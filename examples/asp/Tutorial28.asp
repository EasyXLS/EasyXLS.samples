<%@ Language=VBScript %>

<!-- #INCLUDE FILE="DataType.inc" -->
<%
	'======================================================================
	'Tutorial 28
	'
	' This tutorial shows how to export data to XLS file that has
	' multiple sheets in ASP classic. The first sheet is filled with data.
	'======================================================================
	
	response.write("Tutorial 28<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
	
	' Get the table of data for the first worksheet
	Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()
	
	' Add data in cells for report header
	for column = 0 to 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (column + 1))
		xlsFirstTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	next
	xlsFirstTable.easy_getRowAt(0).setHeight(30)

	' Add data in cells for report values
	for row = 0 to 99
		for column = 0 to 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		next
	next

	' Set column widths
	xlsFirstTable.setColumnWidth_2 0, 70
	xlsFirstTable.setColumnWidth_2 1, 100
	xlsFirstTable.setColumnWidth_2 2, 70
	xlsFirstTable.setColumnWidth_2 3, 100
	xlsFirstTable.setColumnWidth_2 4, 70

	' Export the XLS file
	response.write("Writing file: C:\Samples\Tutorial28 - export XLS file.xls<br>")
	workbook.easy_WriteXLSFile ("C:\Samples\Tutorial28 - export XLS file.xls")
	
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
