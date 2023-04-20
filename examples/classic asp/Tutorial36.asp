<%@ Language=VBScript %>

<%
	'===============================================================================
	' Tutorial 36
	'
	' This tutorial shows how to read an Excel XLSX file in Classic ASP
	' (the XLSX file generated by Tutorial 04 as base template), modify
	' some data and save it to another XLSX file (Tutorial36 - read XLSX file.xlsx).
	'===============================================================================
	
	response.write("Tutorial 36<br>")
	response.write("----------<br>")

	' Create an instance of the class that reads Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Read XLSX file
	response.write("Reading file: C:\Samples\Tutorial04.xlsx<br>")
	if (workbook.easy_LoadXLSXFile("C:\Samples\Tutorial04.xlsx")) then

		' Get the table of data for the second worksheet
		set xlsTable = workbook.easy_getSheet("Second tab").easy_getExcelTable()
		
		' Write some data to the second sheet
		xlsTable.easy_getCell_2("A1").setValue("Data added by Tutorial36")

		for column=0 to 4
			xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
		
		' Export the new XLSX file
		response.write("Writing file: C:\Samples\Tutorial36 - read XLSX file.xlsx<br>")
		workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial36 - read XLSX file.xlsx")
		
		' Confirm export of Excel file
		if workbook.easy_getError() = "" then
			response.write("File successfully created.")
		else
			response.write("Error encountered: " + workbook.easy_getError())
		end if
	else
		response.write("Error reading file C:\Samples\Tutorial04.xlsx")
		response.write(workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>