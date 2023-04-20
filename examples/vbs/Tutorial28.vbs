    '==================================================================
    'Tutorial 28
    '
    ' This tutorial shows how to export data to XLS file  that has
	' multiple sheets in VBScript. The first sheet is filled with data.
    '==================================================================
    
    WScript.StdOut.WriteLine("Tutorial 28" & vbcrlf & "----------" & vbcrlf)
    
	' Constants declaration
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"

    ' Create an instance of the class that exports Excel files
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
	For row = 0 To 99
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

    ' Export the XLS file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial28 - export XLS file.xls")
    workbook.easy_WriteXLSFile ("C:\Samples\Tutorial28 - export XLS file.xls")
    
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
    