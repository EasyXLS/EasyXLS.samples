    '============================================================
    ' Tutorial 20
    '
    ' This tutorial shows how to create an Excel file in VBScript
	' and apply an auto-filter to a range of cells.
    '============================================================
    
    ' Constants declaration
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"

    WScript.StdOut.WriteLine("Tutorial 20" & vbcrlf & "-----------" & vbcrlf)
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create a sheet
	workbook.easy_addWorksheet_2("Sheet1")
	    
	' Get the table of data for the worksheet
    Set xlsTab = workbook.easy_getSheet("Sheet1")
	Set xlsTable = xlsTab.easy_getExcelTable()
	
	' Add data in cells for report header
	For Column = 0 To 4
		xlsTable.easy_getCell(0,column).setValue("Column " & (Column + 1))
		xlsTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	Next
	
	' Add data in cells for report values
	For row = 0 To 99
		For column = 0 To 4
			xlsTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		Next
	Next

	' Apply auto-filter on cell range A1:E1
    Set xlsFilter = xlsTab.easy_getFilter()
    xlsFilter.setAutoFilter_2("A1:E1")

    ' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx")
    
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
    