<%@ Language=VBScript %>

<%
	'=====================================================================
	'Tutorial 11
	'
	' This tutorial shows how to create an Excel file in Classic ASP that
	' has a cell that contains SUM formula for a range of cells.
	'=====================================================================
	
	response.write("Tutorial 11<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create a sheet
	workbook.easy_addWorksheet_2("Formula")
	
	' Get the table of data for the sheet, add data in sheet and the formula
	Set xlsTable = workbook.easy_getSheet("Formula").easy_getExcelTable()
	xlsTable.easy_getCell_2("A1").setValue("1")
	xlsTable.easy_getCell_2("A2").setValue("2")
	xlsTable.easy_getCell_2("A3").setValue("3")
	xlsTable.easy_getCell_2("A4").setValue("4")
	xlsTable.easy_getCell_2("A6").setValue("=SUM(A1:A4)")

	' Export Excel file
	response.write("Writing file: C:\Samples\Tutorial11 - formulas in Excel.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial11 - formulas in Excel.xlsx")
	
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
