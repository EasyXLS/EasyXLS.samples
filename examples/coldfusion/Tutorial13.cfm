<!--
=====================================================================
Tutorial 13

This tutorial shows how to create an Excel file in ColdFusion having
multiple sheets. The second sheet contains a named area range.
The A1:A10 cell range contains data validators, drop down list
and whole number validation.
=====================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataValidator" name="DataValidator" action="CREATE">

	
Tutorial 13<br>
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
<cfset xlsSecondTab.easy_addName("Range", "=Second tab!$A$1:$A$4")>

<!-- Add data validation as drop down list type -->
<cfset xlsFirstTab = workbook.easy_getSheetAt(0)>
<cfset xlsFirstTab.easy_addDataValidator("A1:A10", DataValidator.VALIDATE_LIST, DataValidator.OPERATOR_EQUAL_TO, "=Range", "")>

<!-- Add data validation as whole number type -->
<cfset xlsFirstTab.easy_addDataValidator("B1:B10", DataValidator.VALIDATE_WHOLE_NUMBER, DataValidator.OPERATOR_BETWEEN, "=4", "=100")>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial13 - cell validation in Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial13 - cell validation in Excel.xlsx")>

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