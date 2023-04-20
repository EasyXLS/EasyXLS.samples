    '===================================================================
    ' Tutorial 14
    '
    ' This tutorial shows how to create an Excel file in VBScript having
	' a sheet and conditional formatting for cell ranges.
    '===================================================================
    
    ' Constants declaration
    Dim Bisque, Red
    Bisque = &hffc4e4ff
    Red = &hff0000ff

    Dim DT_NUMERIC
    DT_NUMERIC = "numeric"
    
    Dim CONDITIONAL_FORMATTING_OPERATOR_BETWEEN, CONDITIONAL_FORMATTING_OPERATOR_EQUALTO, CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA
    CONDITIONAL_FORMATTING_OPERATOR_BETWEEN = 1
    CONDITIONAL_FORMATTING_OPERATOR_EQUALTO = 3
    CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA = 2

    WScript.StdOut.WriteLine("Tutorial 14" & vbcrlf & "-----------" & vbcrlf)
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
	' Create a sheet
	workbook.easy_addWorksheet_2("Sheet1")
    
	' Get the table of data for the first worksheet
	Set xlsTab = workbook.easy_getSheet("Sheet1")
	Set xlsTable = xlsTab.easy_getExcelTable()

	' Add data in cells
    For i = 0 To 5
        For j = 0 To 3
			If ((i < 2) And (j < 2)) Then
                xlsTable.easy_getCell(i, j).setValue ("12")
            Else
                If ((j = 2) And (i < 2)) Then
                    xlsTable.easy_getCell(i, j).setValue ("1000")
                Else
                    xlsTable.easy_getCell(i, j).setValue ("9")
                End If
            End If
            xlsTable.easy_getCell(i, j).setDataType (DT_NUMERIC)
        Next
    Next

	' Set conditional formatting
	xlsTab.easy_addConditionalFormatting_5 "A1:C3", CONDITIONAL_FORMATTING_OPERATOR_BETWEEN, "=9", "=11", True, True, Clng(RED)

	' Set another conditional formatting
	xlsTab.easy_addConditionalFormatting_9 "A6:C6", CONDITIONAL_FORMATTING_OPERATOR_BETWEEN, "=COS(PI())+2", "", Clng(BISQUE)
	xlsTab.easy_getConditionalFormattingAt_2("A6:C6").getConditionAt(0).setConditionType(CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA)
    
    ' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial14 - conditional formatting in Excel.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial14 - conditional formatting in Excel.xlsx")
    
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
    