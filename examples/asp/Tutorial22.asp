<%@ Language=VBScript %>

<!-- #INCLUDE FILE="Format.inc" -->
<!-- #INCLUDE FILE="Color.inc" -->
<!-- #INCLUDE FILE="Chart.inc" -->
<%
	'================================================================
	' Tutorial 22
	'
	' This tutorial shows how to create an Excel file in ASP classic
	' with a chart and show and format the chart data table.
	'================================================================
	
	response.write("Tutorial 22<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	set workbook = Server.CreateObject("EasyXLS.ExcelDocument")
	
	' Create a worksheet
	workbook.easy_addWorksheet_2("SourceData")
	
	' Get the table of data for the worksheet
	Set xlsTable1 = workbook.easy_getSheet("SourceData").easy_getExcelTable()

	' Add data in cells for report header
	xlsTable1.easy_getCell(0, 0).setValue("Show Date")
	xlsTable1.easy_getCell(0, 1).setValue("Available Places")
	xlsTable1.easy_getCell(0, 2).setValue("Available Tickets")
	xlsTable1.easy_getCell(0, 3).setValue("Sold Tickets")

	' Add data in cells for chart report values
	xlsTable1.easy_getCell(1, 0).setValue("03/13/2005 00:00:00")
	xlsTable1.easy_getCell(1, 0).setFormat(FORMAT_FORMAT_DATE)
	xlsTable1.easy_getCell(2, 0).setValue("03/14/2005 00:00:00")
	xlsTable1.easy_getCell(2, 0).setFormat(FORMAT_FORMAT_DATE)
	xlsTable1.easy_getCell(3, 0).setValue("03/15/2005 00:00:00")
	xlsTable1.easy_getCell(3, 0).setFormat(FORMAT_FORMAT_DATE)
	xlsTable1.easy_getCell(4, 0).setValue("03/16/2005 00:00:00")
	xlsTable1.easy_getCell(4, 0).setFormat(FORMAT_FORMAT_DATE)
	
	xlsTable1.easy_getCell(1, 1).setValue("10000")
	xlsTable1.easy_getCell(2, 1).setValue("5000")
	xlsTable1.easy_getCell(3, 1).setValue("8500")
	xlsTable1.easy_getCell(4, 1).setValue("1000")

	xlsTable1.easy_getCell(1, 2).setValue("8000")
	xlsTable1.easy_getCell(2, 2).setValue("4000")
	xlsTable1.easy_getCell(3, 2).setValue("6000")
	xlsTable1.easy_getCell(4, 2).setValue("1000")

	xlsTable1.easy_getCell(1, 3).setValue("920")
	xlsTable1.easy_getCell(2, 3).setValue("1005")
	xlsTable1.easy_getCell(3, 3).setValue("342")
	xlsTable1.easy_getCell(4, 3).setValue("967")

	' Set column widths
	xlsTable1.easy_getColumnAt(0).setWidth(100)
	xlsTable1.easy_getColumnAt(1).setWidth(100)
	xlsTable1.easy_getColumnAt(2).setWidth(100)
	xlsTable1.easy_getColumnAt(3).setWidth(100)

	' Add a chart sheet
	workbook.easy_addChart_5 "Chart", "=SourceData!$A$1:$D$5", CHART_SERIES_IN_COLUMNS

	' Get the previously added chart
	Set xlsChartSheet = workbook.easy_getSheetAt(1)
	Set xlsChart = xlsChartSheet.easy_getExcelChart()

	' Hide chart legend
	xlsChart.easy_getLegend().setVisible(false)

	' Show chart data table
	xlsChart.easy_getChartDataTable().setVisible(True)
	xlsChart.easy_getChartDataTable().getFontFormat().setFont("Verdana")
	xlsChart.easy_getChartDataTable().getFontFormat().setFontSize(10)
	xlsChart.easy_getChartDataTable().setHorizontalLines(False)
	xlsChart.easy_getChartDataTable().setLegendKey(True)
	xlsChart.easy_getChartDataTable().getLineColorFormat().setLineColor(CLng(COLOR_BLUE))
	xlsChart.easy_getChartDataTable().setVerticalLines(False)
	
	' Export Excel file
	response.write("Writing file: C:\Samples\Tutorial22 - Excel chart datatable.xlsx<br>")
	workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial22 - Excel chart datatable.xlsx")
	
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Dispose memory
	workbook.Dispose
%>
