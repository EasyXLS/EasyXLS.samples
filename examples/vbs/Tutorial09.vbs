    '==========================================================================
    ' Tutorial 09
    '
    ' This tutorial shows how to create an Excel file in VBScript
	' having multiple sheets. The first sheet is filled with data
	' and the cells are formatted and locked.
	' The column header has comments.
	' The first worksheet has header & footer.
	' The first worksheet has print area, rows to repeat at top, center on page,
	' page orientation, page order, paper size, comments print location,
	' print gridlines option and page breaks.
    '==========================================================================
    
    ' Constants declaration
	Dim DATATYPE_STRING
    DATATYPE_STRING = "string"

    Dim ALIGNMENT_CENTER, ALIGNMENT_BOTTOM, ALIGNMENT_LEFT
    ALIGNMENT_CENTER = "center"
    ALIGNMENT_BOTTOM = "bottom"
    ALIGNMENT_LEFT = "left"
    
    Dim Black, Gray, Yellow, DarkGray, Blue
    Black = &hff000000
    Gray = &hff808080
    Yellow = &hff00ffff
    DarkGray = &hffa9a9a9
    Blue = &hffff0000
    
    Dim BORDER_MEDIUM
    BORDER_MEDIUM = 2
    
    Dim HEADER_POSITION_CENTER, HEADER_POSITION_RIGHT
    HEADER_POSITION_CENTER = "C"
    HEADER_POSITION_RIGHT = "R"
    
    Dim FOOTER_POSITION_CENTER
    FOOTER_POSITION_CENTER = "C"
    
   	Dim PAGESETUP_ORIENTATION_PORTRAIT, PAGESETUP_PAGE_ORDER_DOWN_THEN_OVER, PAGESETUP_PAPER_SIZE_A4, PAGESETUP_COMMENTS_AT_END_OF_SHEET
	PAGESETUP_ORIENTATION_PORTRAIT = "Portrait"
	PAGESETUP_PAGE_ORDER_DOWN_THEN_OVER = 0
	PAGESETUP_PAPER_SIZE_A4 = 9
	PAGESETUP_COMMENTS_AT_END_OF_SHEET = 1
        
    WScript.StdOut.WriteLine("Tutorial 09" & vbcrlf & "----------" & vbcrlf)
   
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
	' Create two sheets
	workbook.easy_addWorksheet_2("First tab")
	workbook.easy_addWorksheet_2("Second tab")
	
	' Protect first sheet
     workbook.easy_getSheetAt(0).setSheetProtected(True)
    
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
					
		' Add comment for report header cells
		xlsFirstTable.easy_getCell(0, column).setComment_2("This is column no " & (column + 1))
	Next
	xlsFirstTable.easy_getRowAt(0).setHeight(30)
	
	' Create a formatting style for cells
	Set xlsStyleData = CreateObject("EasyXLS.ExcelStyle")
	xlsStyleData.setHorizontalAlignment(ALIGNMENT_LEFT)
	xlsStyleData.setForeground(CLng(DARKGRAY))
	xlsStyleData.setWrap(False)
	' Protect cells
	xlsStyleData.setLocked(True)
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
    
	' Add header on center section
	Set xlsFirstTab = workbook.easy_getSheetAt(0)
	xlsFirstTab.easy_getHeaderAt_2(HEADER_POSITION_CENTER).InsertSingleUnderline()
	xlsFirstTab.easy_getHeaderAt_2(HEADER_POSITION_CENTER).InsertFile()
	xlsFirstTab.easy_getHeaderAt_2(HEADER_POSITION_CENTER).InsertValue(" - How to create header and footer")

	' Add header on right section
	xlsFirstTab.easy_getHeaderAt_2(HEADER_POSITION_RIGHT).InsertDate()
	xlsFirstTab.easy_getHeaderAt_2(HEADER_POSITION_RIGHT).InsertValue(" ")
	xlsFirstTab.easy_getHeaderAt_2(HEADER_POSITION_RIGHT).InsertTime()

	' Add footer on center section
	xlsFirstTab.easy_getFooterAt_2(FOOTER_POSITION_CENTER).InsertPage()
	xlsFirstTab.easy_getFooterAt_2(FOOTER_POSITION_CENTER).InsertValue(" of ")
	xlsFirstTab.easy_getFooterAt_2(FOOTER_POSITION_CENTER).InsertPages()
    
    ' Get the object that stores the page setup options for the first sheet
    Set xlsPageSetup = xlsFirstTab.easy_getPageSetup()
    ' Set print area
	xlsPageSetup.easy_setPrintArea_3 ("A1:E101")
    ' Set the rows to repeat at top
	xlsPageSetup.easy_setRowsToRepeatAtTop_3 ("$1:$1")
    ' Set center on page option
	xlsPageSetup.setCenterHorizontally (True)
    ' Set page orientation
	xlsPageSetup.setOrientation (PAGESETUP_ORIENTATION_PORTRAIT)
    ' Set page order
	xlsPageSetup.setPageOrder (PAGESETUP_PAGE_ORDER_DOWN_THEN_OVER)
    ' Set paper size
	xlsPageSetup.setPaperSize (PAGESETUP_PAPER_SIZE_A4)
    ' Set where the comments to be printed
	xlsPageSetup.setPrintComments (PAGESETUP_COMMENTS_AT_END_OF_SHEET)
    ' Set the gridlines to be printed
	xlsPageSetup.setPrintGridlines (True)
    
	' Insert page breaks on rows
	xlsFirstTable.easy_insertPageBreakAtRow (21)
    xlsFirstTable.easy_insertPageBreakAtRow (41)
    xlsFirstTable.easy_insertPageBreakAtRow (61)
    xlsFirstTable.easy_insertPageBreakAtRow (81)
    
	' Set page break preview for the sheet
	xlsFirstTab.setPageBreakPreview (True)

    ' Export Excel file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial09 - Excel page setup.xlsx")
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial09 - Excel page setup.xlsx")
    
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
    