    '===============================================================
    ' Tutorial 03
    '
    ' This tutorial shows how to create an Excel file that has
	' multiple sheets in VBScript. The created Excel file is
	' empty and the next tutorial shows how to add data into sheets.
    '===============================================================
    
    WScript.StdOut.WriteLine("Tutorial 03" & vbcrlf & "----------" & vbcrlf)
    
	' Create an instance of the class that creates Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
    
    ' Create Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial03 - create Excel file.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial03 - create Excel file.xlsx")
    
    ' Confirm the creation of Excel file
    If sError = "" Then
		WScript.StdOut.Write(vbcrlf & "File successfully created.")
    Else
		WScript.StdOut.Write(vbcrlf & "Error: " & workbook.easy_getError())
    End If
    	
	' Dispose memory
	workbook.Dispose
	
	Wscript.StdOut.Write(vbcrlf & "Press Enter to exit ...")
	WScript.StdIn.ReadLine()
    