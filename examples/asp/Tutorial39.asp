<%@ Language=VBScript %>

<%
	'==========================================================================
	' Tutorial 39
	'
	' This tutorial shows how to convert CSV file to Excel in ASP classic. The
	' CSV file generated by Tutorial 30 is imported, some data is modified
	' and after that is exported as Excel file.
	'==========================================================================
	
	response.write("Tutorial 39<br>")
	response.write("----------<br>")

	' Create an instance of the class used to import/export Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Import CSV file
	response.write("Reading file: C:\Samples\Tutorial30.csv<br>")
	if (workbook.easy_LoadCSVFile("C:\Samples\Tutorial30.csv")) then
		
		' Set worksheet name
		workbook.easy_getSheetAt(0).setSheetName("First tab")

		' Add new worksheet and add some data in cells (optional step)
		workbook.easy_addWorksheet_2("Second tab")
		set xlsTable = workbook.easy_getSheetAt(1).easy_getExcelTable()
		xlsTable.easy_getCell_2("A1").setValue("Data added by Tutorial39")

		for column=0 to 4
			xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
		
		' Export Excel file
		response.write("Writing file: C:\Samples\Tutorial39 - convert CSV to Excel.xlsx<br>")
		workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial39 - convert CSV to Excel.xlsx")
		
		' Confirm conversion of CSV to Excel
		if workbook.easy_getError() = "" then
			response.write("File successfully created.")
		else
			response.write("Error encountered: " + workbook.easy_getError())
		end if
	else
		response.write("Error reading file C:\Samples\Tutorial30.csv")
		response.write(workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
