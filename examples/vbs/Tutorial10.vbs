    '======================================================================================
    ' Tutorial 10
    '
    ' This tutorial shows how to export an Excel file with a merged cell range in VBScript.
    '======================================================================================
    
    WScript.StdOut.WriteLine("Tutorial 10" & vbcrlf & "-----------" & vbcrlf)
   
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create a worksheet
	workbook.easy_addWorksheet_2("Sheet1")
	
	' Get the table of data for the worksheet
	Set xlsTable = workbook.easy_getSheet("Sheet1").easy_getExcelTable()
	
	' Merge cells by range
	xlsTable.easy_mergeCells_2("A1:C3")
	    
    ' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial10 - merge cells in Excel.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial10 - merge cells in Excel.xlsx")
    
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