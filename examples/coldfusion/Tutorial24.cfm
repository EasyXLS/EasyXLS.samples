<!--
==================================================================
Tutorial 24

This tutorial shows how to create an Excel file in ColdFusion and
to create a chart in a worksheet.
==================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.Format" name="Format" action="CREATE">

	
Tutorial 24<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Create an worksheet -->
<cfset ret = workbook.easy_addWorksheet("SourceData")>

<!-- Get the table of data for the worksheet -->
<cfset xlsTab1 = workbook.easy_getSheet("SourceData")>
<cfset xlsTable1 = xlsTab1.easy_getExcelTable()>

<!-- Add data in cells for report header -->
<cfset xlsTable1.easy_getCell(0, 0).setValue("Show Date")>
<cfset xlsTable1.easy_getCell(0, 1).setValue("Available Places")>
<cfset xlsTable1.easy_getCell(0, 2).setValue("Available Tickets")>
<cfset xlsTable1.easy_getCell(0, 3).setValue("Sold Tickets")>

<!-- Add data in cells for chart report values -->
<cfset xlsTable1.easy_getCell(1, 0).setValue("03/13/2005 00:00:00")>
<cfset xlsTable1.easy_getCell(1, 0).setFormat(Format.FORMAT_DATE)>
<cfset xlsTable1.easy_getCell(2, 0).setValue("03/14/2005 00:00:00")>
<cfset xlsTable1.easy_getCell(2, 0).setFormat(Format.FORMAT_DATE)>
<cfset xlsTable1.easy_getCell(3, 0).setValue("03/15/2005 00:00:00")>
<cfset xlsTable1.easy_getCell(3, 0).setFormat(Format.FORMAT_DATE)>
<cfset xlsTable1.easy_getCell(4, 0).setValue("03/16/2005 00:00:00")>
<cfset xlsTable1.easy_getCell(4, 0).setFormat(Format.FORMAT_DATE)>

<cfset xlsTable1.easy_getCell(1, 1).setValue("10000")>
<cfset xlsTable1.easy_getCell(2, 1).setValue("5000")>
<cfset xlsTable1.easy_getCell(3, 1).setValue("8500")>
<cfset xlsTable1.easy_getCell(4, 1).setValue("1000")>

<cfset xlsTable1.easy_getCell(1, 2).setValue("8000")>
<cfset xlsTable1.easy_getCell(2, 2).setValue("4000")>
<cfset xlsTable1.easy_getCell(3, 2).setValue("6000")>
<cfset xlsTable1.easy_getCell(4, 2).setValue("1000")>

<cfset xlsTable1.easy_getCell(1, 3).setValue("920")>
<cfset xlsTable1.easy_getCell(2, 3).setValue("1005")>
<cfset xlsTable1.easy_getCell(3, 3).setValue("342")>
<cfset xlsTable1.easy_getCell(4, 3).setValue("967")>

<!-- Set column widths -->
<cfset xlsTable1.easy_getColumnAt(0).setWidth(100)>
<cfset xlsTable1.easy_getColumnAt(1).setWidth(100)>
<cfset xlsTable1.easy_getColumnAt(2).setWidth(100)>
<cfset xlsTable1.easy_getColumnAt(3).setWidth(100)>

<!-- Create a chart -->
<cfobject type="java" class="EasyXLS.Charts.ExcelChart" name="xlsChart" action="CREATE">
<cfset xlsChart.setLeftUpperCorner("A10")>
<cfset xlsChart.setSize(600, 300)>
		
<cfset xlsChart.easy_addSeries("=SourceData!$B$1", "=SourceData!$B$2:$B$5")>
<cfset xlsChart.easy_addSeries("=SourceData!$C$1", "=SourceData!$C$2:$C$5")>
<cfset xlsChart.easy_addSeries("=SourceData!$D$1", "=SourceData!$D$2:$D$5")>
<cfset xlsChart.easy_setCategoryXAxisLabels("=SourceData!$A$2:$A$5")>

<!-- Add the chart to the first worksheet -->
<cfset xlsWorksheet = workbook.easy_getSheet("SourceData")>
<cfset xlsWorksheet.easy_addChart(xlsChart)>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial24 - chart inside worksheet.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial24 - chart inside worksheet.xlsx")>

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