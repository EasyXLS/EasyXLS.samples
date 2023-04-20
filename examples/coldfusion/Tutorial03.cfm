<!--
===============================================================
Tutorial 03

This tutorial shows how to create an Excel file that has
multiple sheets in ColdFusion. The created Excel file is 
empty and the next tutorial shows how to add data into sheets.
===============================================================
-->

	
Tutorial 03<br>
----------<br>

<!-- Create an instance of the class that creates Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two sheets -->
<cfset ret = workbook.easy_addWorksheet("First tab")>
<cfset ret = workbook.easy_addWorksheet("Second tab")>

<!-- Create Excel file -->
Writing file C:\Samples\Tutorial03 - create Excel file.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial03 - create Excel file.xlsx")>

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
