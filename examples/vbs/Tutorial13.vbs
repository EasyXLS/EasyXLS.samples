    '===================================================================
    'Tutorial 13
    '
    ' This tutorial shows how to create an Excel file in VBScript having
	' multiple sheets. The second sheet contains a named range.
	' The A1:A10 cell range contains data validators, drop down list
	' and whole number validation.
    '===================================================================

	' Constants declaration
	Dim VALIDATE_LIST, VALIDATE_WHOLE_NUMBER, OPERATOR_BETWEEN, OPERATOR_EQUAL_TO
    VALIDATE_LIST = 3
    VALIDATE_WHOLE_NUMBER = 1
    OPERATOR_BETWEEN = 0
    OPERATOR_EQUAL_TO = 2
    
    WScript.StdOut.WriteLine("Tutorial 13" & vbcrlf & "----------" & vbcrlf)
    
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
	xlsSecondTab.easy_addName_2 "Range", "=Second tab!$A$1:$A$4"

	' Add data validation as drop down list type
	Set xlsFirstTab = workbook.easy_getSheetAt(0)
	xlsFirstTab.easy_addDataValidator_3 "A1:A10", VALIDATE_LIST, OPERATOR_EQUAL_TO, "=Range", ""

	' Add data validation as whole number type
	xlsFirstTab.easy_addDataValidator_3 "B1:B10", VALIDATE_WHOLE_NUMBER, OPERATOR_BETWEEN, "=4", "=100"
	
    ' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial13 - cell validation in Excel.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial13 - cell validation in Excel.xlsx")
    
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
    