    '=================================================================================
    ' Tutorial 17
    '
    ' This tutorial shows how to create an Excel file with groups on rows in VBScript.
	' The Excel file has two worksheets. The first one is full with data and contains
	' the data groups.
    '=================================================================================
    
    ' Constants declaration
    Dim AUTOFORMAT_EASYXLS1
    AUTOFORMAT_EASYXLS1 = 43
    Dim AUTOFORMAT_EASYXLS2
    AUTOFORMAT_EASYXLS2 = 45
    Dim DATAGROUP_GROUP_BY_ROWS
    DATAGROUP_GROUP_BY_ROWS = 0
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"

    WScript.StdOut.WriteLine("Tutorial 17" & vbcrlf & "-----------" & vbcrlf)
  
    ' Create an instance of the class that exports Excel filess
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
    
	' Get the table of data for the first worksheet
	Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()
	
	' Add data in cells for report header
	For Column = 0 To 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (Column + 1))
		xlsFirstTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	Next
	xlsFirstTable.easy_getRowAt(0).setHeight(30)

	' Add data in cells for report values
	For row = 0 To 24
		For column = 0 To 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		Next
	Next

	' Set column widths
	xlsFirstTable.setColumnWidth_2 0, 70
	xlsFirstTable.setColumnWidth_2 1, 100
	xlsFirstTable.setColumnWidth_2 2, 70
	xlsFirstTable.setColumnWidth_2 3, 100
	xlsFirstTable.setColumnWidth_2 4, 70

    ' Group rows and format A1:E26 cell range
    Set xlsFirstDataGroup = CreateObject("EasyXLS.ExcelDataGroup")
    xlsFirstDataGroup.setRange_2 ("A1:E26")
    xlsFirstDataGroup.setGroupType (DATAGROUP_GROUP_BY_ROWS)
    xlsFirstDataGroup.setCollapsed (False)
    Dim xlsAutoFormat
    Set xlsAutoFormat = CreateObject("EasyXLS.ExcelAutoFormat")
    xlsAutoFormat.InitAs (AUTOFORMAT_EASYXLS1)
    xlsFirstDataGroup.setAutoFormat (xlsAutoFormat)
    workbook.easy_getSheetAt(0).easy_addDataGroup (xlsFirstDataGroup)

    ' Group rows and format A2:E10 cell range, outline level two, inside previous group
    Set xlsSecondDataGroup = CreateObject("EasyXLS.ExcelDataGroup")
    xlsSecondDataGroup.setRange_2 ("A2:E10")
    xlsSecondDataGroup.setGroupType (DATAGROUP_GROUP_BY_ROWS)
    xlsSecondDataGroup.setCollapsed (False)
    Dim xlsAutoFormat2
    Set xlsAutoFormat2 = CreateObject("EasyXLS.ExcelAutoFormat")
    xlsAutoFormat2.InitAs (AUTOFORMAT_EASYXLS2)
    xlsSecondDataGroup.setAutoFormat (xlsAutoFormat2)
    workbook.easy_getSheetAt(0).easy_addDataGroup (xlsSecondDataGroup)

    ' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial17 - group data in Excel.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial17 - group data in Excel.xlsx")
    
    ' Confirm export of Excel file
    Dim sError
    sError = workbook.easy_getError()
    If sError = "" Then
		WScript.StdOut.Write(vbcrlf & "File successfully created.")
    Else
		WScript.StdOut.Write(vbcrlf & "Error: " & sError)
    End If
   
	' Dispose memory
	workbook.Dispose
	
	Wscript.StdOut.Write(vbcrlf & "Press Enter to exit ...")
	WScript.StdIn.ReadLine()
    