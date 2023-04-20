<!--
========================================================================== 
Tutorial 09

This tutorial shows how to create an Excel file in ColdFusion
having multiple sheets. The first sheet is filled with data
and the cells are formatted and locked.
The column header has comments.
The first worksheet has header & footer.
The first worksheet has print area, rows to repeat at top, center on page,
page orientation, page order, paper size, comments print location,
print gridlines option and page breaks.
==========================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Border" name="Border" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Alignment" name="Alignment" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Header" name="Header" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Footer" name="Footer" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.PageSetup" name="PageSetup" action="CREATE">
<cfobject type="java" class="java.awt.Color" name="Color" action="CREATE">

	
Tutorial 09<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two worksheets -->
<cfset ret = workbook.easy_addWorksheet("First tab")>
<cfset ret = workbook.easy_addWorksheet("Second tab")>

<!-- Protect first sheet -->
<cfset workbook.easy_getSheetAt(0).setSheetProtected(true)>

<!-- Get the table of data for the first worksheet -->
<cfset xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()>

<!-- Create the formatting style for the header -->
<cfobject type="java" class="EasyXLS.ExcelStyle" name="xlsStyleHeader" action="CREATE">
<cfset xlsStyleHeader.setFont("Verdana")>
<cfset xlsStyleHeader.setFontSize(8)>
<cfset xlsStyleHeader.setItalic(true)>
<cfset xlsStyleHeader.setBold(true)>
<cfset xlsStyleHeader.setForeground(Color.yellow)>
<cfset xlsStyleHeader.setBackground(Color.black)>
<cfset xlsStyleHeader.setBorderColors(Color.gray, Color.gray, Color.gray, Color.gray)>
<cfset xlsStyleHeader.setBorderStyles(Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM)>
<cfset xlsStyleHeader.setHorizontalAlignment(Alignment.ALIGNMENT_CENTER)>
<cfset xlsStyleHeader.setVerticalAlignment(Alignment.ALIGNMENT_BOTTOM)>
<cfset xlsStyleHeader.setWrap(true)>
<cfset xlsStyleHeader.setDataType(DataType.STRING)>

<!-- Add data in cells for report header -->
<cfloop from="0" to="4" index="column">
		<cfset xlsFirstTable.easy_getCell(0,evaluate(column)).setValue("Column " & evaluate(column + 1))>
		<cfset xlsFirstTable.easy_getCell(0,evaluate(column)).setStyle(xlsStyleHeader)>

		<!-- Add comment for report header cells -->
		<cfset xlsFirstTable.easy_getCell(0, evaluate(column)).setComment("This is column no " & evaluate(column + 1))>
</cfloop>
<cfset ret = xlsFirstTable.easy_getRowAt(0).setHeight(30)>

<!-- Create a formatting style for cells -->
<cfobject type="java" class="EasyXLS.ExcelStyle" name="xlsStyleData" action="CREATE">
<cfset xlsStyleData.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)>
<cfset xlsStyleData.setForeground(Color.lightGray)>
<cfset xlsStyleData.setWrap(false)>
<cfset xlsStyleData.setLocked(true)>
<cfset xlsStyleData.setDataType(DataType.STRING)>

<!-- Add data in cells for report values  -->
<cfloop from="0" to="99" index="row">
	<cfloop from="0" to="4" index="column">
		<cfset xlsFirstTable.easy_getCell(evaluate(row + 1),evaluate(column)).setValue("Data " & evaluate(row + 1) & ", " & evaluate(column + 1))>
		<cfset xlsFirstTable.easy_getCell(evaluate(row + 1),evaluate(column)).setStyle(xlsStyleData)>
	</cfloop>
</cfloop>


<!-- Set column widths -->
<cfset xlsFirstTable.setColumnWidth(0, 70)>
<cfset xlsFirstTable.setColumnWidth(1, 100)>
<cfset xlsFirstTable.setColumnWidth(2, 70)>
<cfset xlsFirstTable.setColumnWidth(3, 100)>
<cfset xlsFirstTable.setColumnWidth(4, 70)>

<!-- Add header on center section -->	
<cfset xlsFirstTab = workbook.easy_getSheetAt(0)>
<cfset xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertSingleUnderline()>
<cfset xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertFile()>
<cfset xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertValue(" - How to create header and footer")>

<!-- Add header on right section -->
<cfset xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertDate()>
<cfset xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertValue(" ")>
<cfset xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertTime()>

<!-- Add footer on center section -->	
<cfset xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertPage()>
<cfset xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertValue(" of ")>
<cfset xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertPages()>


<!-- Get the object that stores the page setup options for the first sheet -->
<cfset xlsPageSetup = xlsFirstTab.easy_getPageSetup()>
<!-- Set print area -->
<cfset xlsPageSetup.easy_setPrintArea("A1:E101")>
<!-- Set the rows to repeat at top -->
<cfset xlsPageSetup.easy_setRowsToRepeatAtTop("$1:$1")>
<!-- Set center on page option -->
<cfset xlsPageSetup.setCenterHorizontally(true)>
<!-- Set page orientation -->
<cfset xlsPageSetup.setOrientation(PageSetup.ORIENTATION_PORTRAIT)>
<!-- Set page order -->
<cfset xlsPageSetup.setPageOrder(PageSetup.PAGE_ORDER_DOWN_THEN_OVER)>
<!-- Set paper size -->
<cfset xlsPageSetup.setPaperSize(PageSetup.PAPER_SIZE_A4)>
<!-- Set where the comments to be printed -->
<cfset xlsPageSetup.setPrintComments(PageSetup.COMMENTS_AT_END_OF_SHEET)>
<!-- Set the gridlines to be printed -->
<cfset xlsPageSetup.setPrintGridlines(true)>

<!-- Insert page breaks on rows -->
<cfset xlsFirstTable.easy_insertPageBreakAtRow(21)>
<cfset xlsFirstTable.easy_insertPageBreakAtRow(41)>
<cfset xlsFirstTable.easy_insertPageBreakAtRow(61)>
<cfset xlsFirstTable.easy_insertPageBreakAtRow(81)>

<!-- Set page break preview for the sheet -->
<cfset xlsFirstTab.setPageBreakPreview(true)>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial09 - Excel page setup.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial09 - Excel page setup.xlsx")>

<!-- Confirm export of Excel file -->
<cfset sError = workbook.easy_getError()>
<CFIF (sError  IS "")>
  <cfoutput>
	File successfully created.
  </cfoutput>
<CFELSE>
  <cfoutput>
	Error encountered:  #sError#
  </cfoutput>
</CFIF>

<!-- Dispose memory -->
<cfset workbook.Dispose()>