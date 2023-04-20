    '=================================================================
    'Tutorial 11
    '
    ' This tutorial shows how to create an Excel file in VBScript that
	' has a cell that contains SUM formula for a range of cells.
    '=================================================================
    
    WScript.StdOut.WriteLine("Tutorial 11" & vbcrlf & "----------" & vbcrlf)
    
	' Create an instance of the class that exports Excel files
	Set workbook = CreateObject("EasyXLS.ExcelDocument")
	
	' Create a sheet
	workbook.easy_addWorksheet_2("Formula")
    
    ' Get the table of data for the sheet, add data in sheet and the formula
    Set xlsTable = workbook.easy_getSheet("Formula").easy_getExcelTable()
    xlsTable.easy_getCell_2("A1").setValue ("1")
    xlsTable.easy_getCell_2("A2").setValue ("2")
    xlsTable.easy_getCell_2("A3").setValue ("3")
    xlsTable.easy_getCell_2("A4").setValue ("4")
    xlsTable.easy_getCell_2("A6").setValue ("=SUM(A1:A4)")

    ' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial11 - formulas in Excel.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial11 - formulas in Excel.xlsx")
    
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