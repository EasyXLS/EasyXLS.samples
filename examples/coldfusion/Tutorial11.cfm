<!--
=================================================================== 
Tutorial 11

This tutorial shows how to create an Excel file in ColdFusion that
has a cell that contains SUM formula for a range of cells.
===================================================================
-->

	
Tutorial 11<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Add a worksheet -->
<cfset ret = workbook.easy_addWorksheet("Formula")>

<!-- Get the table of data for the sheet, add data in sheet and the formula -->
<cfset xlsTable = workbook.easy_getSheet("Formula").easy_getExcelTable()>
<cfset xlsTable.easy_getCell("A1").setValue("1")>
<cfset xlsTable.easy_getCell("A2").setValue("2")>
<cfset xlsTable.easy_getCell("A3").setValue("3")>
<cfset xlsTable.easy_getCell("A4").setValue("4")>
<cfset xlsTable.easy_getCell("A6").setValue("=SUM(A1:A4)")>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial11 - formulas in Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial11 - formulas in Excel.xlsx")>

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