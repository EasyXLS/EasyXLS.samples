    '=======================================================================
	' Tutorial 16
	'
	' This tutorial shows how to create an Excel file with image in VBScript 
	' The Excel file has multiple sheets.
	' The first sheet has an image inserted.
	'=======================================================================
    
    WScript.StdOut.WriteLine("Tutorial 16" & vbcrlf & "----------" & vbcrlf)
    
	' Create an instance of the class that exports Excel files
	Set workbook = CreateObject("EasyXLS.ExcelDocument")
	
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
	
	' Insert image into sheet
	workbook.easy_getSheetAt(0).easy_addImage_5 "C:\\Samples\\EasyXLSLogo.JPG", "A1"
		
    ' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial16 - images in Excel.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial16 - images in Excel.xlsx")
    
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
