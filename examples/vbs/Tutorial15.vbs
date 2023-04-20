    '====================================================================================
    ' Tutorial 15
    '
    ' This tutorial shows how to create an Excel file with hyperlinks in VBScript.
	' EasyXLS supports the following hyperlink types:
	'		1 - hyperlink to URL
	'		2 - hyperlink to file
	'		3 - hyperlink to UNC
	'		4 - hyperlink to cell in the same Excel file
	'		5 - hyperlink to name
	' 
	' The link can be placed on a range of cells.
	'
	' Every type of hyperlink accepts a tool tip description.
	'
	' Every type of hyperlink accepts a text mark. A text mark is a link inside the file.
    '====================================================================================
    
    ' Constants declaration
    Dim HYPERLINKTYPE_URL, HYPERLINKTYPE_FILE, HYPERLINKTYPE_UNC, HYPERLINKTYPE_CELL
    HYPERLINKTYPE_URL = "url"
	HYPERLINKTYPE_FILE = "file"
	HYPERLINKTYPE_UNC = "unc"
	HYPERLINKTYPE_CELL = "cell"
    
    WScript.StdOut.WriteLine("Tutorial 15" & vbcrlf & "-----------" & vbcrlf)
       
    ' Create an instance of the class that exports Excel filess
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
	
	Set xlsTab1 = workbook.easy_getSheetAt(0)
	Set xlsTab2 = workbook.easy_getSheetAt(1)
    
	' Create hyperlink to URL
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_URL, "http://www.euoutsourcing.com", "Link to URL", "B2:E2"

	' Create hyperlink to file
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_FILE, "c:\myfile.xls", "Link to file", "B3"

	' Create hyperlink to UNC
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_UNC, "\\computerName\Folder\file.txt", "Link to UNC", "B4:D4"

	' Create hyperlink to cell on second sheet
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_CELL, "'Second tab'!D3", "Link to CELL", "B5"

	' Create a name on the second sheet
	xlsTab2.easy_addName_2 "Name", "=Second tab!$A$1:$A$4"
	
	' Create hyperlink to name 
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_CELL, "Name", "Link to a name", "B6"
	
	' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial15 - hyperlinks in Excel.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial15 - hyperlinks in Excel.xlsx")
    
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
    