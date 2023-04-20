<%@ Language=VBScript %>

<%
	'===============================================================================
	' Tutorial 27
	'
	' This tutorial shows how to create an Excel file in Classic ASP and
	' encrypt the Excel file by setting the password required for opening the file.
	'===============================================================================

	response.write("Tutorial 27<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create two worksheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")

	' Set the password for protecting the Excel file when the file is open
	workbook.easy_getOptions().setPasswordToOpen("password")

	' Export Excel file
	response.write("Writing file: C:\Samples\Tutorial27 - protect Excel with password and encryption.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial27 - protect Excel with password and encryption.xlsx")
	
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
