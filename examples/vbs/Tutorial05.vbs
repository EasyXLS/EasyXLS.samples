    '========================================================================
    'Tutorial 05
    '
    ' This code sample shows how to export data to Excel file in VBScript and
	' format the cells. The Excel file has multiple worksheets.
	' The first one is filled with data and the cells are formatted
    '========================================================================
    
    ' Constants declaration
	Dim DATATYPE_STRING
    DATATYPE_STRING = "string"
    
    Dim ALIGNMENT_CENTER, ALIGNMENT_BOTTOM, ALIGNMENT_LEFT
    ALIGNMENT_CENTER = "center"
    ALIGNMENT_BOTTOM = "bottom"
    ALIGNMENT_LEFT = "left"
    
    Dim Black, Gray, Yellow, DarkGray
    Black = &hff000000
    Gray = &hff808080
    Yellow = &hff00ffff
    DarkGray = &hffa9a9a9
    
    Dim BORDER_MEDIUM
    BORDER_MEDIUM = 2
    
    WScript.StdOut.WriteLine("Tutorial 05" & vbcrlf & "----------" & vbcrlf)
    
   
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
    
	' Get the table of data for the first worksheet
	Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()
    
	' Create the formatting style for the header
	Set xlsStyleHeader = CreateObject("EasyXLS.ExcelStyle")
	xlsStyleHeader.setFont("Verdana")
	xlsStyleHeader.setFontSize(8)
	xlsStyleHeader.setItalic(True)
	xlsStyleHeader.setBold(True)
	xlsStyleHeader.setForeground(CLng(YELLOW))
	xlsStyleHeader.setBackground(CLng(BLACK))
	xlsStyleHeader.setBorderColors CLng(GRAY), CLng(GRAY), CLng(GRAY), CLng(GRAY)
	xlsStyleHeader.setBorderStyles BORDER_MEDIUM, BORDER_MEDIUM, BORDER_MEDIUM, BORDER_MEDIUM
	xlsStyleHeader.setHorizontalAlignment(ALIGNMENT_CENTER)
	xlsStyleHeader.setVerticalAlignment(ALIGNMENT_BOTTOM)
	xlsStyleHeader.setWrap(True)
	xlsStyleHeader.setDataType(DATATYPE_STRING)

    ' Add data in cells for report header
	For column = 0 To 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (column + 1))
		xlsFirstTable.easy_getCell(0,column).setStyle(xlsStyleHeader)
	Next
	xlsFirstTable.easy_getRowAt(0).setHeight(30)
	
	' Create a formatting style for cells
	Set xlsStyleData = CreateObject("EasyXLS.ExcelStyle")
	xlsStyleData.setHorizontalAlignment(ALIGNMENT_LEFT)
	xlsStyleData.setForeground(CLng(DARKGRAY))
	xlsStyleData.setWrap(False)
	xlsStyleData.setDataType(DATATYPE_STRING)
	
	' Add data in cells for report values
	For row = 0 To 99
		For column = 0 To 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setStyle(xlsStyleData)
		Next
	Next

	' Set column widths
	xlsFirstTable.setColumnWidth_2 0, 70
	xlsFirstTable.setColumnWidth_2 1, 100
	xlsFirstTable.setColumnWidth_2 2, 70
	xlsFirstTable.setColumnWidth_2 3, 100
	xlsFirstTable.setColumnWidth_2 4, 70

    ' Export the Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial05 - format Excel cells.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial05 - format Excel cells.xlsx")
    
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
    