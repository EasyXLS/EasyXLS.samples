    '=================================================================
    ' Tutorial 31
    '
    ' This tutorial shows how to export data to HTML file in VBScript.
    '=================================================================
    
    WScript.StdOut.WriteLine("Tutorial 31" & vbcrlf & "----------" & vbcrlf)
    
	' Constants declaration
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"
    Dim AUTOFORMAT_EASYXLS1
    AUTOFORMAT_EASYXLS1 = 43

    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
	' Create a worksheet
	workbook.easy_addWorksheet_2("First tab")
    
	' Get the table of data for the worksheet
	Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()
	
	' Add data in cells for report header
	For Column = 0 To 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (Column + 1))
		xlsFirstTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	Next

	' Add data in cells for report values
	For row = 0 To 99
		For column = 0 To 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		Next
	Next

	' Create an instance of the class used to format the cells
	Dim xlsAutoFormat 
	Set xlsAutoFormat = CreateObject("EasyXLS.ExcelAutoFormat")
	xlsAutoFormat.InitAs(AUTOFORMAT_EASYXLS1)

	' Apply the predefined format to the cells
	xlsFirstTable.easy_setRangeAutoFormat_2 "A1:E101", xlsAutoFormat

    ' Export HTML file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial31 - export HTML file.html")
    workbook.easy_WriteHTMLFile_3 "C:\Samples\Tutorial31 - export HTML file.html","First tab"
    
    ' Confirm export of HTML file
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
    
