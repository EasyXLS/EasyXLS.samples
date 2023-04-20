"""----------------------------------------------------------
Tutorial 26

This tutorial shows how to create an Excel file in Python and
to create a pivot chart. The pivot chart is added to a
workshet and also to a separate chart sheet.
----------------------------------------------------------"""

import gc

from py4j.java_gateway import JavaGateway
from py4j.java_gateway import java_import 
gateway = JavaGateway()

java_import(gateway.jvm,'EasyXLS.*')
java_import(gateway.jvm,'EasyXLS.Constants.*')
java_import(gateway.jvm,'EasyXLS.PivotTables.*')

print("Tutorial 26\n----------\n")

# Create an instance of the class that exports Excel files, having two sheets
workbook = gateway.jvm.ExcelDocument(2, 1)

# Set the sheet names
workbook.easy_getSheetAt(0).setSheetName("First tab")
workbook.easy_getSheetAt(1).setSheetName("Second tab")
workbook.easy_getSheetAt(2).setSheetName("Pivot chart")

# Get the table of data for the first worksheet
xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()

# Add data in cells for report header
xlsFirstTable.easy_getCell(0,0).setValue("Sale agent")
xlsFirstTable.easy_getCell(0,0).setDataType(gateway.jvm.DataType.STRING)
xlsFirstTable.easy_getCell(0,1).setValue("Sale country")
xlsFirstTable.easy_getCell(0,1).setDataType(gateway.jvm.DataType.STRING)
xlsFirstTable.easy_getCell(0,2).setValue("Month")
xlsFirstTable.easy_getCell(0,2).setDataType(gateway.jvm.DataType.STRING)
xlsFirstTable.easy_getCell(0,3).setValue("Year")
xlsFirstTable.easy_getCell(0,3).setDataType(gateway.jvm.DataType.STRING)
xlsFirstTable.easy_getCell(0,4).setValue("Sale amount")
xlsFirstTable.easy_getCell(0,4).setDataType(gateway.jvm.DataType.STRING)

xlsFirstTable.easy_getRowAt(0).setBold(True)

# Add data in cells for report values - the source for pivot chart
xlsFirstTable.easy_getCell(1,0).setValue("John Down")
xlsFirstTable.easy_getCell(1,1).setValue("USA")
xlsFirstTable.easy_getCell(1,2).setValue("June")
xlsFirstTable.easy_getCell(1,3).setValue("2010")
xlsFirstTable.easy_getCell(1,4).setValue("550")
		
xlsFirstTable.easy_getCell(2,0).setValue("Scott Valey")
xlsFirstTable.easy_getCell(2,1).setValue("United Kingdom")
xlsFirstTable.easy_getCell(2,2).setValue("June")
xlsFirstTable.easy_getCell(2,3).setValue("2010")
xlsFirstTable.easy_getCell(2,4).setValue("2300")
		
xlsFirstTable.easy_getCell(3,0).setValue("John Down")
xlsFirstTable.easy_getCell(3,1).setValue("USA")
xlsFirstTable.easy_getCell(3,2).setValue("July")
xlsFirstTable.easy_getCell(3,3).setValue("2010")
xlsFirstTable.easy_getCell(3,4).setValue("3100")
		
xlsFirstTable.easy_getCell(4,0).setValue("John Down")
xlsFirstTable.easy_getCell(4,1).setValue("USA")
xlsFirstTable.easy_getCell(4,2).setValue("June")
xlsFirstTable.easy_getCell(4,3).setValue("2011")
xlsFirstTable.easy_getCell(4,4).setValue("1050")
			
xlsFirstTable.easy_getCell(5,0).setValue("John Down")
xlsFirstTable.easy_getCell(5,1).setValue("USA")
xlsFirstTable.easy_getCell(5,2).setValue("July")
xlsFirstTable.easy_getCell(5,3).setValue("2011")
xlsFirstTable.easy_getCell(5,4).setValue("2400")
		
xlsFirstTable.easy_getCell(6,0).setValue("Steve Marlowe")
xlsFirstTable.easy_getCell(6,1).setValue("France")
xlsFirstTable.easy_getCell(6,2).setValue("June")
xlsFirstTable.easy_getCell(6,3).setValue("2011")
xlsFirstTable.easy_getCell(6,4).setValue("1200")
		
xlsFirstTable.easy_getCell(7,0).setValue("Scott Valey")
xlsFirstTable.easy_getCell(7,1).setValue("United Kingdom")
xlsFirstTable.easy_getCell(7,2).setValue("June")
xlsFirstTable.easy_getCell(7,3).setValue("2011")
xlsFirstTable.easy_getCell(7,4).setValue("700")
		
xlsFirstTable.easy_getCell(8,0).setValue("Scott Valey")
xlsFirstTable.easy_getCell(8,1).setValue("United Kingdom")
xlsFirstTable.easy_getCell(8,2).setValue("July")
xlsFirstTable.easy_getCell(8,3).setValue("2011")
xlsFirstTable.easy_getCell(8,4).setValue("360")

# Create pivot table
xlsPivotTable = gateway.jvm.ExcelPivotTable()
		
xlsPivotTable.setName("Sales")
xlsPivotTable.setSourceRange("First tab!$A$1:$E$9", workbook)
xlsPivotTable.setLocation("A3:G15")
xlsPivotTable.addFieldToRowLabels("Sale agent")
xlsPivotTable.addFieldToColumnLabels("Year")
xlsPivotTable.addFieldToValues("Sale amount","Sale amount per year", gateway.jvm.PivotTable.SUBTOTAL_SUM) 
xlsPivotTable.addFieldToReportFilter("Sale country")
xlsPivotTable.setOutlineForm()
xlsPivotTable.setStyle(gateway.jvm.PivotTable.PIVOT_STYLE_MEDIUM_9)

# Add the pivot table to the second sheet
workbook.easy_getSheet("Second tab").easy_addPivotTable(xlsPivotTable)

# Create pivot chart
xlsPivotChart1 = gateway.jvm.ExcelPivotChart()
xlsPivotChart1.setSize(600, 300)
xlsPivotChart1.setLeftUpperCorner("A10")
xlsPivotChart1.easy_setChartType(gateway.jvm.Chart.CHART_TYPE_PYRAMID_BAR)
xlsPivotChart1.getChartTitle().setText("Sales")
xlsPivotChart1.setPivotTable(xlsPivotTable)

# Add the pivot chart to the second sheet
workbook.easy_getSheet("Second tab").easy_addPivotChart(xlsPivotChart1)

# Create another pivot chart and add it to the chart sheet
xlsPivotChart2 = gateway.jvm.ExcelPivotChart()
xlsPivotChart2.setSize(970, 630)
xlsPivotChart2.easy_setChartType(gateway.jvm.Chart.CHART_TYPE_PYRAMID_BAR)
xlsPivotChart2.getChartTitle().setText("Sales")
xlsPivotChart2.setPivotTable(xlsPivotTable)
workbook.easy_getSheet("Pivot chart").easy_setExcelChart(xlsPivotChart2)

# Export Excel file
print("Writing file C:\\Samples\\Tutorial26 - pivot chart in Excel.xlsx.")
workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial26 - pivot chart in Excel.xlsx")

# Confirm export of Excel file
sError = workbook.easy_getError()

if sError == "":
    print("\nFile successfully created.\n")
else:
    print("\nError encountered: " + sError + "\n\n")

# Dispose memory
gc.collect()