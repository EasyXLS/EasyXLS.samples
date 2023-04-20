    '===================================================================
    'Tutorial 12
    '
    ' This tutorial shows how to create an Excel file in VBScript having
	' multiple sheets. The second sheet contains a named area range.
    '===================================================================
    
    WScript.StdOut.WriteLine("Tutorial 12" & vbcrlf & "----------" & vbcrlf)
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
    
	' Get the table of data for the second worksheet and populate the worksheet
	Set xlsSecondTab = workbook.easy_getSheetAt(1)
	Set xlsSecondTable = xlsSecondTab.easy_getExcelTable()
	xlsSecondTable.easy_getCell_2("A1").setValue("Range data 1")
	xlsSecondTable.easy_getCell_2("A2").setValue("Range data 2")
	xlsSecondTable.easy_getCell_2("A3").setValue("Range data 3")
	xlsSecondTable.easy_getCell_2("A4").setValue("Range data 4")

	' Create a named area range
	xlsSecondTab.easy_addName_2 "Range", "='Second tab'!$A$1:$A$4"

    ' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial12 - name range in Excel.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial12 - name range in Excel.xlsx")
    
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