<!--
========================================================================
Tutorial 40

This tutorial shows how to convert HTML file to Excel in ColdFusion. The
HTML file generated by Tutorial 31 is imported, some data is modified
and after that is exported as Excel file.
========================================================================
-->

	
Tutorial 40<br>
----------<br>


<!-- Create an instance of the class used to import/export Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Import HTML file -->
Reading file C:\Samples\Tutorial31.html<br><br>

<CFIF (workbook.easy_LoadHTMLFile("C:\Samples\Tutorial31.html")  IS True)>

	<!-- Set worksheet name -->
	<cfset ret = workbook.easy_getSheetAt(0).setSheetName("First tab")>

	<!-- Add new worksheet and add some data in cells (optional step) -->	
	<cfset ret = workbook.easy_addWorksheet("Second tab")>
	<cfset xlsTable = workbook.easy_getSheetAt(1).easy_getExcelTable()>
	<cfset xlsTable.easy_getCell("A1").setValue("Data added by Tutorial40")>
	<cfloop from="0" to="4" index="column">
			<cfset xlsTable.easy_getCell(1,evaluate(column)).setValue("Data " & evaluate(column + 1))>
	</cfloop>

	<!-- Export Excel file -->
	Writing file C:\Samples\Tutorial40 - convert HTML to Excel.xlsx<br>
	<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial40 - convert HTML to Excel.xlsx")>
	
	<!-- Confirm conversion of HTML to Excel -->
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

