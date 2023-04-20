<!--
==================================================================================
Tutorial 17

This tutorial shows how to create an Excel file with groups on rows in ColdFusion.
The Excel file has two worksheets. The first one is full with data and contains
the data groups.

==================================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Styles" name="Styles" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.DataGroup" name="DataGroup" action="CREATE">


Tutorial 17<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two worksheets -->
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
<cfloop from="0" to="24" index="row">
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

<!-- Group rows and format A1:E26 cell range -->
<cfobject type="java" class="EasyXLS.ExcelDataGroup" name="xlsFirstDataGroup" action="CREATE"> 
<cfset xlsFirstDataGroup.setRange("A1:E26")>
<cfset xlsFirstDataGroup.setGroupType(DataGroup.GROUP_BY_ROWS)>
<cfset xlsFirstDataGroup.setCollapsed(false)>

<!-- Create an instance of the class used to format the first group -->
<cfobject type="java" class="EasyXLS.ExcelAutoFormat" name="xlsAutoFormat" action="CREATE"> 
<cfset xlsAutoFormat.InitAs(Styles.AUTOFORMAT_EASYXLS1)>
<cfset xlsFirstDataGroup.setAutoFormat(xlsAutoFormat)>
<cfset workbook.easy_getSheetAt(0).easy_addDataGroup(xlsFirstDataGroup)>

<!-- Group rows and format A2:E10 cell range, outline level two, inside previous group -->
<cfobject type="java" class="EasyXLS.ExcelDataGroup" name="xlsSecondDataGroup" action="CREATE"> 
<cfset xlsSecondDataGroup.setRange("A2:E20")>
<cfset xlsSecondDataGroup.setGroupType(DataGroup.GROUP_BY_ROWS)>
<cfset xlsSecondDataGroup.setCollapsed(false)>

<!-- Create an instance of the class used to format the second group -->
<cfobject type="java" class="EasyXLS.ExcelAutoFormat" name="xlsAutoFormat2" action="CREATE"> 
<cfset xlsAutoFormat2.InitAs(Styles.AUTOFORMAT_EASYXLS2)>
<cfset xlsSecondDataGroup.setAutoFormat(xlsAutoFormat2)>
<cfset workbook.easy_getSheetAt(0).easy_addDataGroup(xlsSecondDataGroup)>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial17 - group data in Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial17 - group data in Excel.xlsx")>

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