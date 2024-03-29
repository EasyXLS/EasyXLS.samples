<!--
===========================================================================
Tutorial 37

This tutorial shows how to read an Excel XLS file in ColdFusion (the
XLS file generated by Tutorial 28 as base template), modify
some data and save it to another XLS file (Tutorial37 - read XLS file.xls).
===========================================================================
-->

	
Tutorial 37<br>
----------<br>


<!-- Create an instance of the class that reads Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Read XLS file -->
Reading file C:\Samples\Tutorial28.xls<br><br>

<CFIF (workbook.easy_LoadXLSFile("C:\Samples\Tutorial28.xls")  IS True)>
	<!-- Get the table of data for the second worksheet -->
	<cfset xlsSecondTable = workbook.easy_getSheet("Second tab").easy_getExcelTable()>
		
	<!-- Write some data to the second sheet -->
	<cfset xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial37")>
	<cfloop from="0" to="4" index="column">
			<cfset xlsSecondTable.easy_getCell(1,evaluate(column)).setValue("Data " & evaluate(column + 1))>
	</cfloop>

	<!-- Export the new XLS file -->
	Writing file C:\Samples\Tutorial37 - read XLS file.xls<br>
	<cfset ret = workbook.easy_WriteXLSFile("C:\Samples\Tutorial37 - read XLS file.xls")>
	
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

<CFELSE>
  <cfoutput>
	Error encountered:  #workbook.easy_getError()#
  </cfoutput>
</CFIF>

<!-- Dispose memory -->
<cfset workbook.Dispose()>


