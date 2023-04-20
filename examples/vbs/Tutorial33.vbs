    '==========================================================================
    ' Tutorial 33
    '
    ' This tutorial shows how to set document properties for Excel file in VBScript,
	' like 'Subject' property for summary information, 'Manager' property for
	' document summary information and a custom property.
    '==========================================================================

	WScript.StdOut.WriteLine("Tutorial 33" & vbcrlf & "-----------" & vbcrlf)

    Dim VT_NUMBER
    VT_NUMBER = 5
	
	' Create an instance of the class that exports Excel files
	Set workbook = CreateObject("EasyXLS.ExcelDocument")

	' Create a worksheet
	workbook.easy_addWorksheet_2("Sheet1")

	' Set the 'Subject' document property
	workbook.getSummaryInformation().setSubject("This is the subject")
	
	' Set the 'Manager' document property
	workbook.getDocumentSummaryInformation().setManager("This is the manager")

	' Set a custom document property
	workbook.getDocumentSummaryInformation().setCustomProperty "PropertyName", VT_NUMBER, "4"

	' Export Excel file
	Wscript.StdOut.WriteLine(vbcrlf & "Writing file: C:\Samples\Tutorial33 - Excel file properties.xlsx")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial33 - Excel file properties.xlsx")

	' Confirm export of Excel file
	Dim sError
	sError = workbook.easy_getError()
	If sError = "" Then
	WScript.StdOut.WriteLine(vbcrlf & "File successfully created.")
	Else
		WScript.StdOut.WriteLine(vbcrlf & "Error: " & sError)
	End If   

	' Dispose memory
	workbook.Dispose
    
    Wscript.StdOut.Write(vbcrlf & "Press Enter to exit ...")
    Wscript.StdIn.ReadLine()
