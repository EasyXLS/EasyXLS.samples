<!--
==============================================================================
Tutorial 41

This tutorial shows how to convert XML spreadsheet to Excel in ColdFusion. The
XML Spreadsheet generated by Tutorial 32 is imported, some data is modified
and after that is exported as Excel file.
==============================================================================
-->

	
Tutorial 41<br>
----------<br>


<!-- Create an instance of the class used to import/export Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Import XML Spreadsheet file -->
Reading file C:\Samples\Tutorial32.xml<br><br>

<CFIF (workbook.easy_LoadXMLSpreadsheetFile("C:\Samples\Tutorial32.xml")  IS True)>

	<!-- Get the table of data from the second sheet and add some data in cells (optional step) -->
	<cfset xlsTable = workbook.easy_getSheetAt(1).easy_getExcelTable()>
	<cfset xlsTable.easy_getCell("A1").setValue("Data added by Tutorial41")>
	<cfloop from="0" to="4" index="column">
			<cfset xlsTable.easy_getCell(1,evaluate(column)).setValue("Data " & evaluate(column + 1))>
	</cfloop>

	<!-- Export Excel file -->
	Writing file C:\Samples\Tutorial41 - convert XML spreadsheet to Excel.xlsx<br>
	<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial41 - convert XML spreadsheet to Excel.xlsx")>
	
	<!-- Confirm conversion of XML Spreadsheet to Excel -->
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

<CFELSE>
  <cfoutput>
	Error encountered:  #workbook.easy_getError()#
  </cfoutput>
</CFIF>

<!-- Dispose memory -->
<cfset workbook.Dispose()>

