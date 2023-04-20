<!--
===================================================================
Tutorial 04

This tutorial shows how to export data to XLSX file that has
multiple sheets in ColdFusion. The first sheet is filled with data.
===================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">

	
Tutorial 04<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two sheets -->
<cfset ret = workbook.easy_addWorksheet("First tab")>
<cfset ret = workbook.easy_addWorksheet("Second tab")>

<!-- Get the table of data for the first worksheet -->
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

<!-- Export the XLSX file -->
Writing file C:\Samples\Tutorial04 - export data to Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial04 - export data to Excel.xlsx")>

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
