<%@ Language=VBScript %>

<!-- #INCLUDE FILE="DataType.inc" -->

<%
	'================================================================
	' Tutorial 20
	'
	' This tutorial shows how to create an Excel file in ASP classic
	' and apply an auto-filter to a range of cells.
	'================================================================
	
	response.write("Tutorial 20<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create a sheet
	workbook.easy_addWorksheet_2("Sheet1")
	
	' Get the table of data for the worksheet
	set xlsTab = workbook.easy_getSheet("Sheet1")
	set xlsTable = xlsTab.easy_getExcelTable()
	
	' Add data in cells for report header
	for column = 0 to 4
		xlsTable.easy_getCell(0,column).setValue("Column " & (column + 1))
		xlsTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	next
	
	' Add data in cells for report values
	for row = 0 to 99
		for column = 0 to 4
			xlsTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		next
	next

	' Apply auto-filter on cell range A1:E1
	set xlsFilter = xlsTab.easy_getFilter()
	xlsFilter.setAutoFilter_2("A1:E1")
	
	' Export Excel file
	response.write("Writing file: C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx")
	
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
