"""-------------------------------------------------------------
Tutorial 23

This tutorial shows how to create an Excel file in Python with a
chart and how to set chart type and formatting properties for 
chart area, plot area, axis, series and legend.
-------------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')
java_import(gateway.jvm,'EasyXLS.Constants.*')
java_import(gateway.jvm,'EasyXLS.Drawings.Formatting.*')
java_import(gateway.jvm,'java.awt.Color')

print("Tutorial 23\n-----------\n")

# Create an instance of the class that exports Excel files
workbook = gateway.jvm.ExcelDocument()

# Create an worksheet
workbook.easy_addWorksheet("SourceData")

# Get the table of data for the worksheet
xlsTable1 = workbook.easy_getSheet("SourceData").easy_getExcelTable()

# Add data in cells for report header
xlsTable1.easy_getCell(0, 0).setValue("Show Date")
xlsTable1.easy_getCell(0, 1).setValue("Available Places")
xlsTable1.easy_getCell(0, 2).setValue("Available Tickets")
xlsTable1.easy_getCell(0, 3).setValue("Sold Tickets")

# Add data in cells for chart report values
xlsTable1.easy_getCell(1, 0).setValue("03/13/2005 00:00:00")
xlsTable1.easy_getCell(1, 0).setFormat(gateway.jvm.Format.FORMAT_DATE)
xlsTable1.easy_getCell(2, 0).setValue("03/14/2005 00:00:00")
xlsTable1.easy_getCell(2, 0).setFormat(gateway.jvm.Format.FORMAT_DATE)
xlsTable1.easy_getCell(3, 0).setValue("03/15/2005 00:00:00")
xlsTable1.easy_getCell(3, 0).setFormat(gateway.jvm.Format.FORMAT_DATE)
xlsTable1.easy_getCell(4, 0).setValue("03/16/2005 00:00:00")
xlsTable1.easy_getCell(4, 0).setFormat(gateway.jvm.Format.FORMAT_DATE)

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

# Set column widths
xlsTable1.easy_getColumnAt(0).setWidth(100)
xlsTable1.easy_getColumnAt(1).setWidth(100)
xlsTable1.easy_getColumnAt(2).setWidth(100)
xlsTable1.easy_getColumnAt(3).setWidth(100)

# Add a chart sheet
workbook.easy_addChart("Chart", "=SourceData!$A$1:$D$5", gateway.jvm.Chart.SERIES_IN_COLUMNS)

# Get the previously added chart
xlsChart = workbook.easy_getSheetAt(1).easy_getExcelChart()

# Set chart type
xlsChart.easy_setChartType(gateway.jvm.Chart.CHART_TYPE_CYLINDER_COLUMN)

# Format chart area
xlsChartArea = xlsChart.easy_getChartArea()
xlsChartArea.getLineColorFormat().setLineColor(gateway.jvm.Color.darkGray)
xlsChartArea.getLineStyleFormat().setDashType(gateway.jvm.LineStyleFormat.DASH_TYPE_SOLID)
xlsChartArea.getLineStyleFormat().setWidth(0.25)

# Format chart plot area
xlsPlotArea =  xlsChart.easy_getPlotArea()
xlsPlotArea.getLineColorFormat().setLineColor(gateway.jvm.Color.darkGray)
xlsPlotArea.getLineStyleFormat().setDashType(gateway.jvm.LineStyleFormat.DASH_TYPE_SOLID)
xlsPlotArea.getLineStyleFormat().setWidth(0.25)

# Format chart legend
xlsChartLegend = xlsChart.easy_getLegend()
xlsChartLegend.getFillFormat().setBackground(gateway.jvm.Color.pink)
xlsChartLegend.getFontFormat().setForeground(gateway.jvm.Color.blue)
xlsChartLegend.getFontFormat().setItalic(True)
xlsChartLegend.setKeysArrangementDirection(gateway.jvm.Chart.KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL)
xlsChartLegend.setPlacement(gateway.jvm.Chart.LEGEND_CORNER)
xlsChartLegend.getShadowFormat().setShadow(gateway.jvm.ShadowFormat.OFFSET_DIAGONAL_BOTTOM_RIGHT)

# Format chart X axis
xlsXAxis = xlsChart.easy_getCategoryXAxis()
xlsXAxis.getLineColorFormat().setLineColor(gateway.jvm.Color.lightGray)
xlsXAxis.getLineStyleFormat().setDashType(gateway.jvm.LineStyleFormat.DASH_TYPE_DASH_DOT)
xlsXAxis.getLineStyleFormat().setWidth(0.25)
xlsXAxis.getFontFormat().setForeground(gateway.jvm.Color.red)

# Format chart Y axis
xlsYAxis = xlsChart.easy_getValueYAxis()
xlsYAxis.getLineColorFormat().setLineColor(gateway.jvm.Color.lightGray)
xlsYAxis.getLineStyleFormat().setDashType(gateway.jvm.LineStyleFormat.DASH_TYPE_LONG_DASH)
xlsYAxis.getLineStyleFormat().setWidth(0.25)
xlsYAxis.getFontFormat().setForeground(gateway.jvm.Color.blue)

# Fomat chart series 
xlsChart.easy_getSeriesAt(0).getFillFormat().setBackground(gateway.jvm.Color.blue)
xlsChart.easy_getSeriesAt(1).getFillFormat().setBackground(gateway.jvm.Color.yellow)
xlsChart.easy_getSeriesAt(2).getFillFormat().setBackground(gateway.jvm.Color.green)

# Export Excel file
print("Writing file C:\\Samples\\Tutorial23 - various Excel chart settings.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial23 - various Excel chart settings.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()