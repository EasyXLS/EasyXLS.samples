<!--
================================================================== 
Tutorial 25

This tutorial shows how to create an Excel file in ColdFusion and
to create a pivot table in a worksheet.
==================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.DataType" name="DataType" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.PivotTable" name="PivotTable" action="CREATE">


	
Tutorial 25<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create two worksheets -->
<cfset ret = workbook.easy_addWorksheet("First tab")>
<cfset ret = workbook.easy_addWorksheet("Second tab")>

<!-- Get the table of data for the first worksheet -->
<cfset xlsFirstTab = workbook.easy_getSheet("First tab")>
<cfset xlsFirstTable = xlsFirstTab.easy_getExcelTable()>

<!-- Add data in cells for report header -->
<cfset xlsFirstTable.easy_getCell(0,0).setValue("Sale agent")>
<cfset xlsFirstTable.easy_getCell(0,0).setDataType(DataType.STRING)>
<cfset xlsFirstTable.easy_getCell(0,1).setValue("Sale country")>
<cfset xlsFirstTable.easy_getCell(0,1).setDataType(DataType.STRING)>
<cfset xlsFirstTable.easy_getCell(0,2).setValue("Month")>
<cfset xlsFirstTable.easy_getCell(0,2).setDataType(DataType.STRING)>
<cfset xlsFirstTable.easy_getCell(0,3).setValue("Year")>
<cfset xlsFirstTable.easy_getCell(0,3).setDataType(DataType.STRING)>
<cfset xlsFirstTable.easy_getCell(0,4).setValue("Sale amount")>
<cfset xlsFirstTable.easy_getCell(0,4).setDataType(DataType.STRING)>

<cfset xlsFirstTable.easy_getRowAt(0).setBold(true)>

<!-- Add data in cells for report values - the source for pivot table -->
<cfset xlsFirstTable.easy_getCell(1,0).setValue("John Down")>
<cfset xlsFirstTable.easy_getCell(1,1).setValue("USA")>
<cfset xlsFirstTable.easy_getCell(1,2).setValue("June")>
<cfset xlsFirstTable.easy_getCell(1,3).setValue("2010")>
<cfset xlsFirstTable.easy_getCell(1,4).setValue("550")>

<cfset xlsFirstTable.easy_getCell(2,0).setValue("Scott Valey")>
<cfset xlsFirstTable.easy_getCell(2,1).setValue("United Kingdom")>
<cfset xlsFirstTable.easy_getCell(2,2).setValue("June")>
<cfset xlsFirstTable.easy_getCell(2,3).setValue("2010")>
<cfset xlsFirstTable.easy_getCell(2,4).setValue("2300")>

<cfset xlsFirstTable.easy_getCell(3,0).setValue("John Down")>
<cfset xlsFirstTable.easy_getCell(3,1).setValue("USA")>
<cfset xlsFirstTable.easy_getCell(3,2).setValue("July")>
<cfset xlsFirstTable.easy_getCell(3,3).setValue("2010")>
<cfset xlsFirstTable.easy_getCell(3,4).setValue("3100")>

<cfset xlsFirstTable.easy_getCell(4,0).setValue("John Down")>
<cfset xlsFirstTable.easy_getCell(4,1).setValue("USA")>
<cfset xlsFirstTable.easy_getCell(4,2).setValue("June")>
<cfset xlsFirstTable.easy_getCell(4,3).setValue("2011")>
<cfset xlsFirstTable.easy_getCell(4,4).setValue("1050")>

<cfset xlsFirstTable.easy_getCell(5,0).setValue("John Down")>
<cfset xlsFirstTable.easy_getCell(5,1).setValue("USA")>
<cfset xlsFirstTable.easy_getCell(5,2).setValue("July")>
<cfset xlsFirstTable.easy_getCell(5,3).setValue("2011")>
<cfset xlsFirstTable.easy_getCell(5,4).setValue("2400")>

<cfset xlsFirstTable.easy_getCell(6,0).setValue("Steve Marlowe")>
<cfset xlsFirstTable.easy_getCell(6,1).setValue("France")>
<cfset xlsFirstTable.easy_getCell(6,2).setValue("June")>
<cfset xlsFirstTable.easy_getCell(6,3).setValue("2011")>
<cfset xlsFirstTable.easy_getCell(6,4).setValue("1200")>

<cfset xlsFirstTable.easy_getCell(7,0).setValue("Scott Valey")>
<cfset xlsFirstTable.easy_getCell(7,1).setValue("United Kingdom")>
<cfset xlsFirstTable.easy_getCell(7,2).setValue("June")>
<cfset xlsFirstTable.easy_getCell(7,3).setValue("2011")>
<cfset xlsFirstTable.easy_getCell(7,4).setValue("700")>

<cfset xlsFirstTable.easy_getCell(8,0).setValue("Scott Valey")>
<cfset xlsFirstTable.easy_getCell(8,1).setValue("United Kingdom")>
<cfset xlsFirstTable.easy_getCell(8,2).setValue("July")>
<cfset xlsFirstTable.easy_getCell(8,3).setValue("2011")>
<cfset xlsFirstTable.easy_getCell(8,4).setValue("360")>

<!-- Create pivot table -->
<cfobject type="java" class="EasyXLS.PivotTables.ExcelPivotTable" name="xlsPivotTable" action="CREATE">

<cfset xlsPivotTable.setName("Sales")>
<cfset xlsPivotTable.setSourceRange("First tab!$A$1:$E$9", workbook)>
<cfset xlsPivotTable.setLocation("A3:G15")>
<cfset xlsPivotTable.addFieldToRowLabels("Sale agent")>
<cfset xlsPivotTable.addFieldToColumnLabels("Year")>
<cfset xlsPivotTable.addFieldToValues("Sale amount","Sale amount per year",PivotTable.SUBTOTAL_SUM)>
<cfset xlsPivotTable.addFieldToReportFilter("Sale country")>
<cfset xlsPivotTable.setOutlineForm()>
<cfset xlsPivotTable.setStyle(PivotTable.PIVOT_STYLE_DARK_11)>

<!-- Add pivot table to the second sheet -->
<cfset xlsWorksheet = workbook.easy_getSheet("Second tab")>
<cfset xlsWorksheet.easy_addPivotTable(xlsPivotTable)>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial25 - pivot table in Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial25 - pivot table in Excel.xlsx")>

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