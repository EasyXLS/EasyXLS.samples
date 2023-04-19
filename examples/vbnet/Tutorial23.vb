'-------------------------------------------------------------------
' Tutorial 23
'
' This tutorial shows how to create an Excel file in VB.NET with a
' chart and how to set chart type and formatting properties for
' chart area, plot area, axis, series and legend.
'-------------------------------------------------------------------

Imports System.Drawing
Imports EasyXLS
Imports EasyXLS.Charts
Imports EasyXLS.Drawings.Formatting
Imports EasyXLS.Constants



Module Tutorial23

    Sub Main()

        Console.WriteLine("Tutorial 23" & vbCrLf & "----------" & vbCrLf)

        ' Create an instance of the class that exports Excel files
        Dim workbook As New ExcelDocument

        ' Create an worksheet
        workbook.easy_addWorksheet("SourceData")

        ' Get the table of data for the worksheet
        Dim xlsFirstTab As ExcelWorksheet = workbook.easy_getSheet("SourceData")
        Dim xlsTable1 = xlsFirstTab.easy_getExcelTable()

        ' Add data in cells for report header
        xlsTable1.easy_getCell(0, 0).setValue("Show Date")
        xlsTable1.easy_getCell(0, 1).setValue("Available Places")
        xlsTable1.easy_getCell(0, 2).setValue("Available Tickets")
        xlsTable1.easy_getCell(0, 3).setValue("Sold Tickets")

        ' Add data in cells for chart report values
        xlsTable1.easy_getCell(1, 0).setValue("03/13/2005 00:00:00")
        xlsTable1.easy_getCell(1, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE)
        xlsTable1.easy_getCell(2, 0).setValue("03/14/2005 00:00:00")
        xlsTable1.easy_getCell(2, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE)
        xlsTable1.easy_getCell(3, 0).setValue("03/15/2005 00:00:00")
        xlsTable1.easy_getCell(3, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE)
        xlsTable1.easy_getCell(4, 0).setValue("03/16/2005 00:00:00")
        xlsTable1.easy_getCell(4, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE)

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
        workbook.easy_addChart("Chart", "=SourceData!$A$1:$D$5", Chart.SERIES_IN_COLUMNS)

        ' Get the previously added chart
        Dim xlsChartSheet As ExcelChartSheet = workbook.easy_getSheetAt(1)
        Dim xlsChart As ExcelChart = xlsChartSheet.easy_getExcelChart()

        ' Set chart type
        xlsChart.easy_setChartType(Chart.CHART_TYPE_CYLINDER_COLUMN)

        ' Format chart area
        Dim xlsChartArea = xlsChart.easy_getChartArea()
        xlsChartArea.getLineColorFormat().setLineColor(Color.DarkGray)
        xlsChartArea.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_SOLID)
        xlsChartArea.getLineStyleFormat().setWidth(0.25F)

        ' Format chart plot area
        Dim xlsPlotArea = xlsChart.easy_getPlotArea()
        xlsPlotArea.getLineColorFormat().setLineColor(Color.DarkGray)
        xlsPlotArea.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_SOLID)
        xlsPlotArea.getLineStyleFormat().setWidth(0.25F)

        ' Format chart legend
        Dim xlsChatLegend = xlsChart.easy_getLegend()
        xlsChatLegend.getFillFormat().setBackground(Color.LavenderBlush)
        xlsChatLegend.getFontFormat().setForeground(Color.Blue)
        xlsChatLegend.getFontFormat().setItalic(True)
        xlsChatLegend.setKeysArrangementDirection(Chart.KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL)
        xlsChatLegend.setPlacement(Chart.LEGEND_CORNER)
        xlsChatLegend.getShadowFormat().setShadow(ShadowFormat.OFFSET_DIAGONAL_BOTTOM_RIGHT)

        ' Format chart X axis
        Dim xlsXAxis = xlsChart.easy_getCategoryXAxis()
        xlsXAxis.getLineColorFormat().setLineColor(Color.SteelBlue)
        xlsXAxis.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_DASH_DOT)
        xlsXAxis.getLineStyleFormat().setWidth(0.25F)
        xlsXAxis.getFontFormat().setForeground(Color.Red)

        ' Format chart Y axis
        Dim xlsYAxis = xlsChart.easy_getValueYAxis()
        xlsYAxis.getLineColorFormat().setLineColor(Color.SteelBlue)
        xlsYAxis.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_LONG_DASH)
        xlsYAxis.getLineStyleFormat().setWidth(0.25F)
        xlsYAxis.getFontFormat().setForeground(Color.Blue)

        ' Fomat chart series
        xlsChart.easy_getSeriesAt(0).getFillFormat().setBackground(Color.RoyalBlue)
        xlsChart.easy_getSeriesAt(1).getFillFormat().setBackground(Color.Yellow)
        xlsChart.easy_getSeriesAt(2).getFillFormat().setBackground(Color.LightGreen)

        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial23 - various Excel chart settings.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial23 - various Excel chart settings.xlsx")

        ' Confirm export of Excel file
        Dim sError As String = workbook.easy_getError()
        If (sError.Equals("")) Then
            Console.Write(vbCrLf & "File successfully created. Press Enter to Exit...")
        Else
            Console.Write(vbCrLf & "Error encountered: " & sError & vbCrLf & "Press Enter to Exit...")
        End If

        ' Dispose memory
        workbook.Dispose()

        Console.ReadLine()
    End Sub

End Module
