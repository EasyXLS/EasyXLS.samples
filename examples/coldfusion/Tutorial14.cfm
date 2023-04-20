<!--
==============================================================
Tutorial 14

This tutorial shows how to create an Excel file in ColdFusion 
having a sheet and conditional formatting for cell ranges.
==============================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.ConditionalFormatting" name="ConditionalFormatting" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.FontSettings" name="FontSettings" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Border" name="Border" action="CREATE">
<cfobject type="java" class="java.awt.Color" name="Color" action="CREATE">

	
Tutorial 14<br>
----------<br>


<!-- Create an instance of the class that exports Excel -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Add a worksheet -->
<cfset ret = workbook.easy_addWorksheet("Sheet1")>

<!-- Get the table of data for the first worksheet -->
<cfset xlsTab = workbook.easy_getSheet("Sheet1")>
<cfset xlsTable = xlsTab.easy_getExcelTable()>

<!-- Add data in cells -->
<cfloop from="0" to="5" index="i">
	<cfloop from="0" to="3" index="j">
		<cfif ((i LT 2) AND (j LT 2))>
			<cfset xlsTable.easy_getCell(evaluate(i), evaluate(j)).setValue("12")>
		<cfelseif ((j EQ 2) AND (i LT 2))>
			<cfset xlsTable.easy_getCell(evaluate(i), evaluate(j)).setValue("1000")>
		<cfelse>
			<cfset xlsTable.easy_getCell(evaluate(i), evaluate(j)).setValue("9")>
		</cfif>
		<cfset xlsTable.easy_getCell(evaluate(i), evaluate(j)).setDataType(DataType.NUMERIC)>
	</cfloop>
</cfloop>

<!-- Set conditional formatting -->
<cfset xlsTab.easy_addConditionalFormatting("A1:C3", ConditionalFormatting.OPERATOR_BETWEEN, "=9", "=11", true, true, Color.RED)>

<!-- Set another conditional formatting -->
<cfset xlsTab.easy_addConditionalFormatting("A6:C6", ConditionalFormatting.OPERATOR_BETWEEN, "=COS(PI())+2", "", Color.ORANGE)>
<cfset xlsTab.easy_getConditionalFormattingAt("A6:C6").getConditionAt(0).setConditionType(ConditionalFormatting.CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA)>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial14 - conditional formatting in Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial14 - conditional formatting in Excel.xlsx")>

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