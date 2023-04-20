<!--
==================================================================
Tutorial 31

This tutorial shows how to export data to HTML file in ColdFusion.
==================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Styles" name="Styles" action="CREATE">


Tutorial 31<br>
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

<!-- Apply a predefined format to the cells -->
<cfobject type="java" class="EasyXLS.ExcelAutoFormat" name="xlsAutoFormat" action="CREATE"> 
<cfset xlsAutoFormat.InitAs(Styles.AUTOFORMAT_EASYXLS1)>
<cfset xlsFirstTable.easy_setRangeAutoFormat("A1:E101", xlsAutoFormat)>


<!-- Export HTML file -->
Writing file C:\Samples\Tutorial31 - export HTML file.html<br>
<cfset ret = workbook.easy_WriteHTMLFile("C:\Samples\Tutorial31 - export HTML file.html", "First tab")>

<!-- Confirm export of HTML file -->
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
