<%@ Language=VBScript %>

<!-- #INCLUDE FILE="DataType.inc" -->
<%
	'====================================================================
	'Tutorial 30
	'
	' This tutorial shows how to export data to CSV file in Classic ASP.
	'====================================================================
	
	response.write("Tutorial 30<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create a worksheet
	workbook.easy_addWorksheet_2("First tab")
	
	' Get the table of data for the worksheet
	Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()
	
	' Add data in cells for report header
	for column = 0 to 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (column + 1))
		xlsFirstTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	next

	' Add data in cells for report values
	for row = 0 to 99
		for column = 0 to 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		next
	next

	' Export CSV file
	response.write("Writing file: C:\Samples\Tutorial30 - export CSV file.csv<br>")
	workbook.easy_WriteCSVFile "C:\Samples\Tutorial30 - export CSV file.csv", "First tab"
	
	' Confirm export of CSV file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
