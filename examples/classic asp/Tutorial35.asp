<%@ Language=VBScript %>

<%
	'=======================================================================
	' Tutorial 35
	'
	' This tutorial shows how to import Excel sheet to List in Classic ASP.
	' The data is imported from a specific Excel sheet (For this example
	' we use the Excel file generated in Tutorial 09).
	'=======================================================================
	
	response.write("Tutorial 35<br>")
	response.write("----------<br>")

	' Create an instance of the class that imports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Import Excel sheet to List
	response.write("Reading file: C:\Samples\Tutorial09.xlsx<br><br>")
	Set rows = workbook.easy_ReadXLSXSheet_AsList_3("C:\Samples\Tutorial09.xlsx", "First tab")

	' Confirm import of Excel file
    if workbook.easy_getError() = "" then
		' Display imported List values
		for rowIndex = 0 to rows.size() - 1
			Set row = rows.elementAt(rowIndex)
			for cellIndex = 0 to row.size - 1
				response.write("At row " & (rowIndex + 1) & ", column " & (cellIndex + 1) & " the value is '" & row.elementAt(cellIndex) & "'<br>")
			next
		next
	else
		response.Write("Error reading file C:\Samples\Tutorial09.xlsx " & workbook.easy_getError())
    end if

	' Dispose memory
	workbook.Dispose
%>

