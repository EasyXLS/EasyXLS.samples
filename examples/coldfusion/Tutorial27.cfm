<!--
=============================================================================
Tutorial 27

This tutorial shows how to create an Excel file in ColdFusion and
encrypt the Excel file by setting the password required for opening the file.
=============================================================================
-->

	
Tutorial 27<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two worksheets -->
<cfset ret = workbook.easy_addWorksheet("First tab")>
<cfset ret = workbook.easy_addWorksheet("Second tab")>

<!-- Set the password for protecting the Excel file when the file is open -->
<cfset ret = workbook.easy_getOptions().setPasswordToOpen("password")>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial27 - protect Excel with password and encryption.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial27 - protect Excel with password and encryption.xlsx")>

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
