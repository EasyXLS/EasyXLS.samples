<!--
===================================================================== 
Tutorial 12

This tutorial shows how to create an Excel file in ColdFusion having
multiple sheets. The second sheet contains a named area range.
=====================================================================
-->

	
Tutorial 12<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two worksheets -->
<cfset ret = workbook.easy_addWorksheet("First tab")>
<cfset ret = workbook.easy_addWorksheet("Second tab")>

<!-- Get the table of data for the second worksheet and populate the worksheet -->
<cfset xlsSecondTab = workbook.easy_getSheetAt(1)>
<cfset xlsSecondTable = xlsSecondTab.easy_getExcelTable()>
<cfset xlsSecondTable.easy_getCell("A1").setValue("Range data 1")>
<cfset xlsSecondTable.easy_getCell("A2").setValue("Range data 2")>
<cfset xlsSecondTable.easy_getCell("A3").setValue("Range data 3")>
<cfset xlsSecondTable.easy_getCell("A4").setValue("Range data 4")>

<!-- Create a named area range -->
<cfset xlsSecondTab.easy_addName("Range", "='Second tab'!$A$1:$A$4")>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial12 - name range in Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial12 - name range in Excel.xlsx")>

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