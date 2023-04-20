<%@ Language=VBScript %>

<!-- #INCLUDE FILE="HyperlinkType.inc" -->
<%
	'======================================================================================
	' Tutorial 15
	'
	' This tutorial shows how to create an Excel file with hyperlinks in ASP classic.
	'
	' EasyXLS supports the following hyperlink types:
	'		1 - hyperlink to URL 
	'		2 - hyperlink to file 
	'		3 - hyperlink to UNC 
	'		4 - hyperlink to cell in the same Excel file
	'		5 - hyperlink to name 
	'
	' The link can be placed on a range of cells.
	'
	' Every type of hyperlink accepts a tool tip description.
	'
	' Every type of hyperlink accepts a text mark. A text mark is a link inside the file.
	'======================================================================================
	
	response.write("Tutorial 15<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
	
	set xlsTab1 = workbook.easy_getSheetAt(0)
	set xlsTab2 = workbook.easy_getSheetAt(1)
	
	' Create hyperlink to URL 
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_URL, "http://www.euoutsourcing.com", "Link to URL", "B2:E2"

	' Create hyperlink to file
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_FILE, "c:\myfile.xls", "Link to file", "B3"

	' Create hyperlink to UNC
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_UNC, "\\computerName\Folder\file.txt", "Link to UNC", "B4:D4"

	' Create hyperlink to cell on second sheet
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_CELL, "'Second tab'!D3", "Link to CELL", "B5"

	' Create a name on the second sheet
	xlsTab2.easy_addName_2 "Name", "=Second tab!$A$1:$A$4"
	
	' Create hyperlink to name
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_CELL, "Name", "Link to a name", "B6"
	
	' Export Excel file
	response.write("Writing file: C:\Samples\Tutorial15 - hyperlinks in Excel.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial15 - hyperlinks in Excel.xlsx")
	
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>