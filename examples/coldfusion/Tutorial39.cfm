<!--
=======================================================================
Tutorial 39

This tutorial shows how to convert CSV file to Excel in ColdFusion. The
CSV file generated by Tutorial 30 is imported, some data is modified
and after that is exported as Excel file.
=======================================================================
-->

	
Tutorial 39<br>
----------<br>


<!-- Create an instance of the class used to import/export Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Import CSV file -->
Reading file C:\Samples\Tutorial30.csv<br><br>

<CFIF (workbook.easy_LoadCSVFile("C:\Samples\Tutorial30.csv")  IS True)>

	<!-- Set worksheet name -->
	<cfset ret = workbook.easy_getSheetAt(0).setSheetName("First tab")>

	<!-- Add new worksheet and add some data in cells (optional step) -->	
	<cfset ret = workbook.easy_addWorksheet("Second tab")>
	<cfset xlsTable = workbook.easy_getSheetAt(1).easy_getExcelTable()>
	<cfset xlsTable.easy_getCell("A1").setValue("Data added by Tutorial39")>
	<cfloop from="0" to="4" index="column">
			<cfset xlsTable.easy_getCell(1,evaluate(column)).setValue("Data " & evaluate(column + 1))>
	</cfloop>

	<!-- Export Excel file -->
	Writing file C:\Samples\Tutorial39 - convert CSV to Excel.xlsx<br>
	<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial39 - convert CSV to Excel.xlsx")>
	
	<!-- Confirm conversion of CSV to Excel -->
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


