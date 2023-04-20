<!--
=====================================================================
Tutorial 19

This tutorial shows how to create an Excel file in ColdFusion having
multiple sheets. The first sheet is filled with data and the
first cell of the second row contains data in rich text format.

=====================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">

	
Tutorial 19<br>
----------<br>


<!-- Create an instance of the class that exports Excel -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two worksheets -->
<cfset ret = workbook.easy_addWorksheet("First tab")>
<cfset ret = workbook.easy_addWorksheet("Second tab")>

<!-- Get the table of data for the first worksheet -->
<cfset xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()>

<!-- Create the string used to set the RTF in cell -->
<cfset sFormattedValue = "This is <b>bold</b>.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "This is <i>italic</i>.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "This is <u>underline</u>.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "This is <underline double>double underline</underline double>.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "This is <font color=red>red</font>.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "This is <font color=rgb(255,0,0)>red</font> too.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "This is <font face=""Arial Black"">Arial Black</font>.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "This is <font size=15pt>size 15</font>.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "This is <s>strikethrough</s>.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "This is <sup>superscript</sup>.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "This is <sub>subscript</sub>.">
<cfset sFormattedValue = sFormattedValue & Chr(10) & "<b>This</b> <i>is</i> <font color=red face=""Arial Black"" size=15pt><underline double>formatted</underline double></font> <s>text</s>.">

<!-- Set the rich text value in cell -->
<cfset xlsFirstTable.easy_getCell(1, 0).setHTMLValue(sFormattedValue)>
<cfset xlsFirstTable.easy_getCell(1, 0).setDataType(DataType.STRING)>
<cfset xlsFirstTable.easy_getCell(1, 0).setWrap(true)>
<cfset xlsFirstTable.easy_getRowAt(1).setHeight(250)>
<cfset xlsFirstTable.easy_getColumnAt(0).setWidth(250)>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial19 - RTF for Excel cells.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial19 - RTF for Excel cells.xlsx")>

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