<%@ Language=VBScript %>

<!-- #INCLUDE FILE="DataValidator.inc" -->
<%
	'=======================================================================
	'Tutorial 13
	'
	' This tutorial shows how to create an Excel file in Classic ASP having
	' multiple sheets. The second sheet contains a named range.
	' The A1:A10 cell range contains data validators, drop down list
	' and whole number validation.
	'=======================================================================
	
	response.write("Tutorial 13<br>")
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
	xlsSecondTab.easy_addName_2 "Range", "=Second tab!$A$1:$A$4"
	
	' Add data validation as drop down list type
	set xlsFirstTab = workbook.easy_getSheetAt(0)
	xlsFirstTab.easy_addDataValidator_3 "A1:A10", DATAVALIDATOR_VALIDATE_LIST, DATAVALIDATOR_OPERATOR_EQUAL_TO, "=Range", ""

	' Add data validation as whole number type
	xlsFirstTab.easy_addDataValidator_3 "B1:B10", DATAVALIDATOR_VALIDATE_WHOLE_NUMBER, DATAVALIDATOR_OPERATOR_BETWEEN, "=4", "=100"

	' Export Excel file
	response.write("Writing file: C:\Samples\Tutorial13 - cell validation in Excel.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial13 - cell validation in Excel.xlsx")
		
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
