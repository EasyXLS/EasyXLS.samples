<!--
=======================================================================================
Tutorial 10

This tutorial shows how to export an Excel file with a merged cell range in ColdFusion.
=======================================================================================
-->

	
Tutorial 10<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Add a worksheet -->
<cfset ret = workbook.easy_addWorksheet("Sheet1")>

<!-- Get the table of data for the worksheet -->
<cfset xlsTable = workbook.easy_getSheet("Sheet1").easy_getExcelTable()>
	
<!-- Merge cells by range -->
<cfset xlsTable.easy_mergeCells("A1:C3")>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial10 - merge cells in Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial10 - merge cells in Excel.xlsx")>

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