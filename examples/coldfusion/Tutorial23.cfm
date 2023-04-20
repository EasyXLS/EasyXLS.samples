<!--
===================================================================
Tutorial 23

This tutorial shows how to create an Excel file in ColdFusion with
a chart and how to set chart type and formatting properties for
chart area, plot area, axis, series and legend.
===================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.Format" name="Format" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Chart" name="Chart" action="CREATE">
<cfobject type="java" class="EasyXLS.Drawings.Formatting.LineStyleFormat" name="LineStyleFormat" action="CREATE">
<cfobject type="java" class="EasyXLS.Drawings.Formatting.ShadowFormat" name="ShadowFormat" action="CREATE">
<cfobject type="java" class="java.awt.Color" name="Color" action="CREATE">

	
Tutorial 23<br>
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

<!-- Add a chart sheet -->
<cfset workbook.easy_addChart("Chart", "=SourceData!$A$1:$D$5", Chart.SERIES_IN_COLUMNS)>

<!-- Get the previously added chart -->
<cfset xlsChartSheet = workbook.easy_getSheetAt(1)>
<cfset xlsChart = xlsChartSheet.easy_getExcelChart()>

<!-- Set chart type -->
<cfset xlsChart.easy_setChartType(Chart.CHART_TYPE_CYLINDER_COLUMN)>

<!-- Format chart area -->
<cfset xlsChartArea = xlsChart.easy_getChartArea()>
<cfset xlsChartArea.getLineColorFormat().setLineColor(Color.DarkGray)>
<cfset xlsChartArea.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_SOLID)>
<cfset xlsChartArea.getLineStyleFormat().setWidth(0.25)>
		
<!-- Format chart plot area -->
<cfset xlsPlotArea =  xlsChart.easy_getPlotArea()>
<cfset xlsPlotArea.getLineColorFormat().setLineColor(Color.DarkGray)>
<cfset xlsPlotArea.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_SOLID)>
<cfset xlsPlotArea.getLineStyleFormat().setWidth(0.25)>

<!-- Format chart legend -->
<cfset xlsChartLegend = xlsChart.easy_getLegend()>
<cfset xlsChartLegend.getFillFormat().setBackground(Color.Pink)>
<cfset xlsChartLegend.getFontFormat().setForeground(Color.Blue)>
<cfset xlsChartLegend.getFontFormat().setItalic(true)>
<cfset xlsChartLegend.setKeysArrangementDirection(Chart.KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL)>
<cfset xlsChartLegend.setPlacement(Chart.LEGEND_CORNER)>
<cfset xlsChartLegend.getShadowFormat().setShadow(ShadowFormat.OFFSET_DIAGONAL_BOTTOM_RIGHT)>

<!-- Format chart X axis -->
<cfset xlsXAxis = xlsChart.easy_getCategoryXAxis()>
<cfset xlsXAxis.getLineColorFormat().setLineColor(Color.LightGray)>
<cfset xlsXAxis.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_DASH_DOT)>
<cfset xlsXAxis.getLineStyleFormat().setWidth(0.25)>
<cfset xlsXAxis.getFontFormat().setForeground(Color.Red)>

<!-- Format chart Y axis -->
<cfset xlsYAxis = xlsChart.easy_getValueYAxis()>
<cfset xlsYAxis.getLineColorFormat().setLineColor(Color.LightGray)>
<cfset xlsYAxis.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_LONG_DASH)>
<cfset xlsYAxis.getLineStyleFormat().setWidth(0.25)>
<cfset xlsYAxis.getFontFormat().setForeground(Color.Blue)>

<!-- Fomat chart series -->
<cfset xlsChart.easy_getSeriesAt(0).getFillFormat().setBackground(Color.Blue)>
<cfset xlsChart.easy_getSeriesAt(1).getFillFormat().setBackground(Color.Yellow)>
<cfset xlsChart.easy_getSeriesAt(2).getFillFormat().setBackground(Color.Green)>

<!-- Export Excel file -->
Writing file C:\Samples\Tutorial23 - various Excel chart settings.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile("C:\Samples\Tutorial23 - various Excel chart settings.xlsx")>

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