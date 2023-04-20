<!--
=================================================================
Tutorial 30

This tutorial shows how to export data to CSV file in ColdFusion.
=================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">

	
Tutorial 30<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create a worksheet -->
<cfset ret = workbook.easy_addWorksheet("First tab")>

<!-- Get the table of data for the worksheet -->
<cfset xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()>

<!-- Add data in cells for report header -->
<cfloop from="0" to="4" index="column">
		<cfset xlsFirstTable.easy_getCell(0,evaluate(column)).setValue("Column " & evaluate(column + 1))>
		<cfset xlsFirstTable.easy_getCell(0,evaluate(column)).setDataType(DataType.STRING)>
</cfloop>
	
<!-- Add data in cells for report values  -->
<cfloop from="0" to="99" index="row">
	<cfloop from="0" to="4" index="column">
		<cfset xlsFirstTable.easy_getCell(evaluate(row + 1),evaluate(column)).setValue("Data " & evaluate(row + 1) & ", " & evaluate(column + 1))>
		<cfset xlsFirstTable.easy_getCell(evaluate(row + 1),evaluate(column)).setDataType(DataType.STRING)>
	</cfloop>
</cfloop>

<!-- Export CSV file -->
Writing file C:\Samples\Tutorial30 - export CSV file.csv<br>
<cfset ret = workbook.easy_WriteCSVFile("C:\Samples\Tutorial30 - export CSV file.csv", "First tab")>

<!-- Confirm export of CSV file -->
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
