<%@ Language=VBScript %>

<%
	'=======================================================================
	'Tutorial 12
	'
	' This tutorial shows how to create an Excel file in Classic ASP having
	' multiple sheets. The second sheet contains a named area range.
	'=======================================================================
	
	response.write("Tutorial 12<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
	
	' Get the table of data for the second worksheet and populate the worksheet
	set xlsSecondTab = workbook.easy_getSheetAt(1)
	set xlsSecondTable = xlsSecondTab.easy_getExcelTable()
	xlsSecondTable.easy_getCell_2("A1").setValue("Range data 1")
	xlsSecondTable.easy_getCell_2("A2").setValue("Range data 2")
	xlsSecondTable.easy_getCell_2("A3").setValue("Range data 3")
	xlsSecondTable.easy_getCell_2("A4").setValue("Range data 4")

	' Create a named area range
	xlsSecondTab.easy_addName_2 "Range", "='Second tab'!$A$1:$A$4"

	' Export Excel file
	response.write("Writing file: C:\Samples\Tutorial12 - name range in Excel.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial12 - name range in Excel.xlsx")
	
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
