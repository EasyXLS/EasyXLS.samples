<!--
====================================================================
Tutorial 28

This tutorial shows how to export data to XLS file that has
multiple sheets in ColdFusion. The first sheet is filled with data.
====================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">

	
Tutorial 28<br>
----------<br>


<!-- Create an instance of the class that exports Excel files, having two sheets -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Set the sheet names -->
<cfset ret = workbook.easy_addWorksheet("First tab")>
<cfset ret = workbook.easy_addWorksheet("Second tab")>

<!-- Get the table of data for the first worksheet -->
<cfset xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()>

<!-- Add data in cells for report header -->
<cfloop from="0" to="4" index="column">
		<cfset xlsFirstTable.easy_getCell(0,evaluate(column)).setValue("Column " & evaluate(column + 1))>
		<cfset xlsFirstTable.easy_getCell(0,evaluate(column)).setDataType(DataType.STRING)>
</cfloop>
<cfset ret = xlsFirstTable.easy_getRowAt(0).setHeight(30)>
	
<!-- Add data in cells for report values  -->
<cfloop from="0" to="99" index="row">
	<cfloop from="0" to="4" index="column">
		<cfset xlsFirstTable.easy_getCell(evaluate(row + 1),evaluate(column)).setValue("Data " & evaluate(row + 1) & ", " & evaluate(column + 1))>
		<cfset xlsFirstTable.easy_getCell(evaluate(row + 1),evaluate(column)).setDataType(DataType.STRING)>
	</cfloop>
</cfloop>

<!-- Set column widths -->
<cfset xlsFirstTable.setColumnWidth(0, 70)>
<cfset xlsFirstTable.setColumnWidth(1, 100)>
<cfset xlsFirstTable.setColumnWidth(2, 70)>
<cfset xlsFirstTable.setColumnWidth(3, 100)>
<cfset xlsFirstTable.setColumnWidth(4, 70)>

<!-- Export the XLS file -->
Writing file C:\Samples\Tutorial28 - export XLS file.xls<br>
<cfset ret = workbook.easy_WriteXLSFile("C:\Samples\Tutorial28 - export XLS file.xls")>

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
