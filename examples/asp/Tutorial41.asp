<%@ Language=VBScript %>

<%
	'===============================================================
	' Tutorial 41
	'
	' This tutorial shows how to load an XML file (we use the file
	' generated in Tutorial 32), modify some data and save it to
	' another file (Tutorial41.xls).
	'===============================================================
	
	response.write("Tutorial 41<br>")
	response.write("----------<br>")


	' Create an instance of the object that generates Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Read the file
	response.write("Reading file: C:\Samples\Tutorial32.xml<br>")
	if (workbook.easy_LoadXMLSpreadsheetFile_2("C:\Samples\Tutorial32.xml")) then
		
		' Get the table of the second worksheet and write some data
		set xlsTable = workbook.easy_getSheetAt(1).easy_getExcelTable()
		xlsTable.easy_getCell_2("A1").setValue("Data added by Tutorial41")

		for column=0 to 4
			xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
		
		' Generate the file
		response.write("Writing file: C:\Samples\Tutorial41 - convert XML spreadsheet to Excel.xlsx<br>")
		workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial41 - convert XML spreadsheet to Excel.xlsx")
		
		' Confirm generation
		if workbook.easy_getError() = "" then
			response.write("File successfully created.")
		else
			response.write("Error encountered: " + workbook.easy_getError())
		end if
	else
		response.write("Error reading file C:\Samples\Tutorial32.xml")
		response.write(workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
