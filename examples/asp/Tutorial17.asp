<%@ Language=VBScript %>

<!-- #INCLUDE FILE="DataType.inc" -->
<!-- #INCLUDE FILE="Styles.inc" -->
<!-- #INCLUDE FILE="DataGroup.inc" -->

<%
	'======================================================================================
	' Tutorial 17
	'
    ' This tutorial shows how to create an Excel file with groups on rows in ASP classic.
	' The Excel file has two worksheets. The first one is full with data and contains the
	' data groups.
	'======================================================================================
	
	response.write("Tutorial 17<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
	
	' Get the table of data for the first worksheet
	Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()
	
	' Add data in cells for report header
	for column = 0 to 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (column + 1))
		xlsFirstTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	next
	xlsFirstTable.easy_getRowAt(0).setHeight(30)

	' Add data in cells for report values
	for row = 0 to 24
		for column = 0 to 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		next
	next

	' Set column widths
	xlsFirstTable.setColumnWidth_2 0, 70
	xlsFirstTable.setColumnWidth_2 1, 100
	xlsFirstTable.setColumnWidth_2 2, 70
	xlsFirstTable.setColumnWidth_2 3, 100
	xlsFirstTable.setColumnWidth_2 4, 70

	' Group rows and format A1:E26 cell range
    Set xlsFirstDataGroup = Server.CreateObject("EasyXLS.ExcelDataGroup")
    xlsFirstDataGroup.setRange_2 ("A1:E26")
    xlsFirstDataGroup.setGroupType (DATAGROUP_GROUP_BY_ROWS)
    xlsFirstDataGroup.setCollapsed (False)
    Dim xlsAutoFormat
    Set xlsAutoFormat = Server.CreateObject("EasyXLS.ExcelAutoFormat")
    xlsAutoFormat.InitAs (AUTOFORMAT_EASYXLS1)
    xlsFirstDataGroup.setAutoFormat (xlsAutoFormat)
    workbook.easy_getSheetAt(0).easy_addDataGroup (xlsFirstDataGroup)

    ' Group rows and format A2:E10 cell range, outline level two, inside previous group
    Set xlsSecondDataGroup = Server.CreateObject("EasyXLS.ExcelDataGroup")
    xlsSecondDataGroup.setRange_2 ("A2:E10")
    xlsSecondDataGroup.setGroupType (DATAGROUP_GROUP_BY_ROWS)
    xlsSecondDataGroup.setCollapsed (False)
    Dim xlsAutoFormat2
    Set xlsAutoFormat2 = Server.CreateObject("EasyXLS.ExcelAutoFormat")
    xlsAutoFormat2.InitAs (AUTOFORMAT_EASYXLS2)
    xlsSecondDataGroup.setAutoFormat (xlsAutoFormat2)
    workbook.easy_getSheetAt(0).easy_addDataGroup (xlsSecondDataGroup)

	' Export Excel file
	response.write("Writing file: C:\Samples\Tutorial17 - group data in Excel.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial17 - group data in Excel.xlsx")
	
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
