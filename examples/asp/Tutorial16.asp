<%@ Language=VBScript %>

<%
	'===========================================================================
	'Tutorial 16
	'
	' This tutorial shows how to create an Excel file with image in ASP classic
	' The Excel file has multiple sheets.
	' The first sheet has an image inserted.
	'===========================================================================
	
	response.write("Tutorial 16<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
	
	' Insert image into sheet
	workbook.easy_getSheetAt(0).easy_addImage_5 "C:\\Samples\\EasyXLSLogo.JPG", "A1"

	' Export Excel file
	response.write("Writing file: C:\Samples\Tutorial16 - images in Excel.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial16 - images in Excel.xlsx")
	
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
