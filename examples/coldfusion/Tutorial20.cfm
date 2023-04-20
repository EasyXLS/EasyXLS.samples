<!--
==============================================================
Tutorial 20

This tutorial shows how to create an Excel file in ColdFusion
and apply an auto-filter to a range of cells.

==============================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">

	
Tutorial 20<br>
----------<br>


<!-- Create an instance of the class that exports Excel -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two worksheets -->
<cfset ret = workbook.easy_addWorksheet("Sheet1")>

<!-- Get the table of data for the worksheet -->
<cfset xlsTab = workbook.easy_getSheet("Sheet1")>
<cfset xlsTable = xlsTab.easy_getExcelTable()>

<!-- Add data in cells for report header -->
<cfloop from="0" to="4" index="column">
		<cfset xlsTable.easy_getCell(0,evaluate(column)).setValue("Column " & evaluate(column + 1))>
		<cfset xlsTable.easy_getCell(0,evaluate(column)).setDataType(DataType.STRING)>
</cfloop>
	
<!-- Add data in cells for report values -->
<cfloop from="0" to="99" index="row">
	<cfloop from="0" to="4" index="column">
		<cfset xlsTable.easy_getCell(evaluate(row + 1),evaluate(column)).setValue("Data " & evaluate(row + 1) & ", " & evaluate(column + 1))>
		<cfset xlsTable.easy_getCell(evaluate(row + 1),evaluate(column)).setDataType(DataType.STRING)>
	</cfloop>
</cfloop>

<!-- Apply auto-filter on cell range A1:E1 -->
<cfset xlsFilter = xlsTab.easy_getFilter()>
<cfset xlsFilter.setAutoFilter("A1:E1")>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx")>

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