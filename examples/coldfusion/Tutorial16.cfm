<!--
=========================================================================
Tutorial 16

This tutorial shows how to create an Excel file with image in ColdFusion 
The Excel file has multiple sheets.
The first worksheet has an image inserted.
=========================================================================
-->

	
Tutorial 16<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two sheets -->
<cfset ret = workbook.easy_addWorksheet("First tab")>
<cfset ret = workbook.easy_addWorksheet("Second tab")>

<!-- Insert image into sheet -->
<cfset workbook.easy_getSheetAt(0).easy_addImage("C:\Samples\EasyXLSLogo.JPG", "A1")>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial16 - images in Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial16 - images in Excel.xlsx")>

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