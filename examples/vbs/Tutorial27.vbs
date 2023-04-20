    '==============================================================================
    ' Tutorial 27
    '
    ' This tutorial shows how to create an Excel file in VBScript and
	' encrypt the Excel file by setting the password required for opening the file.
    '==============================================================================
    
    WScript.StdOut.WriteLine("Tutorial 27" & vbcrlf & "----------" & vbcrlf)

    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
	' Create two worksheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
	
	' Set the password for protecting the Excel file when the file is open
	workbook.easy_getOptions().setPasswordToOpen("password")
    
    ' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial27 - protect Excel with password and encryption.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial27 - protect Excel with password and encryption.xlsx")
    
    ' Confirm export of Excel file
    If sError = "" Then
		WScript.StdOut.Write(vbcrlf & "File successfully created.")
    Else
		WScript.StdOut.Write(vbcrlf & "Error: " & workbook.easy_getError())
    End If
    	
	' Dispose memory
	workbook.Dispose
	
	Wscript.StdOut.Write(vbcrlf & "Press Enter to exit ...")
	WScript.StdIn.ReadLine()
