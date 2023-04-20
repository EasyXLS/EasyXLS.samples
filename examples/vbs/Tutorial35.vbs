    '===================================================================
    ' Tutorial 35
    '
    ' This tutorial shows how to import Excel sheet to List in VBScript.
	' The data is imported from a specific Excel sheet (For this example
	' we use the Excel file generated in Tutorial 09).
    '===================================================================
    
    WScript.StdOut.WriteLine("Tutorial 35" & vbcrlf & "-----------" & vbcrlf)
    
	' Create an instance of the class that imports Excel files
	Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Import Excel sheet to List
    WScript.StdOut.WriteLine("Reading file: C:\Samples\Tutorial09.xlsx")
    WScript.StdOut.WriteLine()
    Set rows = workbook.easy_ReadXLSXSheet_AsList_3("C:\Samples\Tutorial09.xlsx", "First tab")

	' Confirm import of Excel file
    If workbook.easy_getError() = "" Then
		' Display imported List values
		For rowIndex = 0 To rows.Size() - 1
			Set row = rows.elementAt(rowIndex)
			For cellIndex = 0 To row.Size - 1
				WScript.StdOut.WriteLine("At row " & (rowIndex + 1) & ", column " & (cellIndex + 1) & " the value is '" & row.elementAt(cellIndex) & "'")
			Next
		Next
	Else
		WScript.StdOut.Write(vbcrlf & "Error reading file C:\Samples\Tutorial09.xls " & workbook.easy_getError())
    End If
    
    ' Dispose memory
	workbook.Dispose
	
    Wscript.StdOut.Write("Press Enter to exit ...")
    Wscript.StdIn.ReadLine
