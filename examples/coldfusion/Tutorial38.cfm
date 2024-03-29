<!--
==============================================================================
Tutorial 38

This tutorial shows how to read an Excel XLSB file in ColdFusion (the
XLSB file generated by Tutorial 29 as base template), modify
some data and save it to another XLSB file (Tutorial38 - read XLSB file.xlsb).
==============================================================================
-->

	
Tutorial 38<br>
----------<br>


<!-- Create an instance of the class that reads Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Read XLSB file -->
Reading file C:\Samples\Tutorial29.xlsb<br><br>

<CFIF (workbook.easy_LoadXLSBFile("C:\Samples\Tutorial29.xlsb")  IS True)>
	<!-- Get the table of data for the second worksheet -->
	<cfset xlsTable = workbook.easy_getSheetAt(1).easy_getExcelTable()>
	
	<!-- Write some data to the second sheet -->
	<cfset xlsTable.easy_getCell("A1").setValue("Data added by Tutorial38")>
	<cfloop from="0" to="4" index="column">
			<cfset xlsTable.easy_getCell(1,evaluate(column)).setValue("Data " & evaluate(column + 1))>
	</cfloop>

	<!-- Export the new XLSB file -->
	Writing file C:\Samples\Tutorial38 - read XLSB file.xlsb<br>
	<cfset ret = workbook.easy_WriteXLSBFile("C:\Samples\Tutorial38 - read XLSB file.xlsb")>
	
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


