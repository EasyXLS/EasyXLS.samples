<%@ Language=VBScript %>

<%
	'==============================================================
	' Tutorial 03
	'
	' This tutorial shows how to create an Excel file that has
	' multiple sheets in Classic ASP. The created Excel file is 
	' empty and the next tutorial shows how to add data into sheets.
	'==============================================================

	response.write("Tutorial 03<br>")
	response.write("----------<br>")

	' Create an instance of the class that creates Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")

	' Create Excel file
	response.write("Writing file: C:\Samples\Tutorial03 - create Excel file.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial03 - create Excel file.xlsx")
	
	' Confirm the creation of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
