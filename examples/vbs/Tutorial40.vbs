    '=======================================================================
    ' Tutorial 40
    '
    ' This tutorial shows how to convert HTML file to Excel in VBScript. The
	' HTML file generated by Tutorial 31 is imported, some data is modified
	' and after that is exported as Excel file.
    '=======================================================================
    
    WScript.StdOut.WriteLine("Tutorial 40" & vbcrlf & "-----------" & vbcrlf)

	' Create an instance of the class used to import/export Excel files
	Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Import HTML file
    WScript.StdOut.WriteLine("Reading file: C:\Samples\Tutorial31.html" & vbcrlf)
    If (workbook.easy_LoadHTMLFile_2("C:\Samples\Tutorial31.html")) Then
    
		' Set worksheet name
		workbook.easy_getSheetAt(0).setSheetName("First tab")

		' Add new worksheet and add some data in cells (optional step)
		workbook.easy_addWorksheet_2("Second tab")
		Set xlsTable = workbook.easy_getSheetAt(1).easy_getExcelTable()

		xlsTable.easy_getCell_2("A1").setValue("Data added by Tutorial40")

		For column=0 To 4
			xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		Next
        
        ' Export Excel file
        Wscript.StdOut.WriteLine(vbcrlf & "Writing file: C:\Samples\Tutorial40 - convert HTML to Excel.xlsx")
        workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial40 - convert HTML to Excel.xlsx")
    
		' Confirm conversion of HTML to Excel
		Dim sError
		sError = workbook.easy_getError()
		If sError = "" Then
		WScript.StdOut.WriteLine(vbcrlf & "File successfully created.")
		Else
			WScript.StdOut.WriteLine(vbcrlf & "Error: " & sError)
		End If    
    Else
        Wscript.StdOut.WriteLine("Error reading file C:\Samples\Tutorial31.html" & vbcrlf & workbook.easy_getError())
    End If
    
	' Dispose memory
	workbook.Dispose
    
    Wscript.StdOut.Write(vbcrlf & "Press Enter to exit ...")
    Wscript.StdIn.ReadLine()
