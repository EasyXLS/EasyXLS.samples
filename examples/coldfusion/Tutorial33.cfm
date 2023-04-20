<!--
================================================================================
Tutorial 33

This tutorial shows how to set document properties for Excel file in ColdFusion,
like 'Subject' property for summary information, 'Manager' property for
document summary information and a custom property.
================================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.FileProperty" name="FileProperty" action="CREATE">
	
	
Tutorial 33<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Add a worksheet -->
<cfset ret = workbook.easy_addWorksheet("Sheet1")>

<!-- Set the 'Subject' document property -->
<cfset workbook.getSummaryInformation().setSubject("This is the subject")>
	
<!-- Set the 'Manager' document property -->
<cfset workbook.getDocumentSummaryInformation().setManager("This is the manager")>

<!-- Set a custom document property -->
<cfset workbook.getDocumentSummaryInformation().setCustomProperty("PropertyName", FileProperty.VT_NUMBER, "4")>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial33 - Excel file properties.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial33 - Excel file properties.xlsx")>

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
