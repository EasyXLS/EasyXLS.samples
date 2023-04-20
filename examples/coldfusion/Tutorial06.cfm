<!--
======================================================================
Tutorial 06

This code sample shows how to create an Excel file in ColdFusion with
multiple sheets. The first sheet is protected and
filled with data. The cells are formatted and locked.
======================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Border" name="Border" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Alignment" name="Alignment" action="CREATE">
<cfobject type="java" class="java.awt.Color" name="Color" action="CREATE">

	
Tutorial 06<br>
----------<br>


<!-- Create an instance of the class that creates Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two sheets -->
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

<!-- Create Excel file -->
Writing file C:\Samples\Tutorial06 - protect Excel sheet.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial06 - protect Excel sheet.xlsx")>

<!-- Confirm the creation of Excel file -->
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
