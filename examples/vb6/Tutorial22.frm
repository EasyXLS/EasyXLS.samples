VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   100
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    '=======================================================
	' Tutorial 22
	'
	' This tutorial shows how to create an Excel file in VB6
	' with a chart and show and format the chart data table.
	'=======================================================
    
    Format.Initialize
    Color.Initialize
    Chart.Initialize
        
    Me.Label1.Caption = "Tutorial 22" & vbCrLf & "-----------------" & vbCrLf

    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create a worksheet
    workbook.easy_addWorksheet_2 ("SourceData")
    
    ' Get the table of data for the worksheet
    Set xlsTable1 = workbook.easy_getSheet("SourceData").easy_getExcelTable()

	' Add data in cells for report header
    xlsTable1.easy_getCell(0, 0).setValue ("Show Date")
    xlsTable1.easy_getCell(0, 1).setValue ("Available Places")
    xlsTable1.easy_getCell(0, 2).setValue ("Available Tickets")
    xlsTable1.easy_getCell(0, 3).setValue ("Sold Tickets")

	' Add data in cells for chart report values
    xlsTable1.easy_getCell(1, 0).setValue ("03/13/2005 00:00:00")
    xlsTable1.easy_getCell(1, 0).setFormat (Format.FORMAT_FORMAT_DATE)
    xlsTable1.easy_getCell(2, 0).setValue ("03/14/2005 00:00:00")
    xlsTable1.easy_getCell(2, 0).setFormat (Format.FORMAT_FORMAT_DATE)
    xlsTable1.easy_getCell(3, 0).setValue ("03/15/2005 00:00:00")
    xlsTable1.easy_getCell(3, 0).setFormat (Format.FORMAT_FORMAT_DATE)
    xlsTable1.easy_getCell(4, 0).setValue ("03/16/2005 00:00:00")
    xlsTable1.easy_getCell(4, 0).setFormat (Format.FORMAT_FORMAT_DATE)
    
    xlsTable1.easy_getCell(1, 1).setValue ("10000")
    xlsTable1.easy_getCell(2, 1).setValue ("5000")
    xlsTable1.easy_getCell(3, 1).setValue ("8500")
    xlsTable1.easy_getCell(4, 1).setValue ("1000")

    xlsTable1.easy_getCell(1, 2).setValue ("8000")
    xlsTable1.easy_getCell(2, 2).setValue ("4000")
    xlsTable1.easy_getCell(3, 2).setValue ("6000")
    xlsTable1.easy_getCell(4, 2).setValue ("1000")

    xlsTable1.easy_getCell(1, 3).setValue ("920")
    xlsTable1.easy_getCell(2, 3).setValue ("1005")
    xlsTable1.easy_getCell(3, 3).setValue ("342")
    xlsTable1.easy_getCell(4, 3).setValue ("967")

	' Set column widths
    xlsTable1.easy_getColumnAt(0).setWidth (100)
    xlsTable1.easy_getColumnAt(1).setWidth (100)
    xlsTable1.easy_getColumnAt(2).setWidth (100)
    xlsTable1.easy_getColumnAt(3).setWidth (100)

    ' Add a chart sheet
    workbook.easy_addChart_5 "Chart", "=SourceData!$A$1:$D$5", Chart.CHART_SERIES_IN_COLUMNS

    ' Get the previously added chart
    Set xlsChartSheet = workbook.easy_getSheetAt(1)
    Set xlsChart = xlsChartSheet.easy_getExcelChart()
    
    ' Hide chart legend
    xlsChart.easy_getLegend().setVisible (False)

    ' Show chart data table
    xlsChart.easy_getChartDataTable().setVisible (True)
    xlsChart.easy_getChartDataTable().getFontFormat().setFont ("Verdana")
    xlsChart.easy_getChartDataTable().getFontFormat().setFontSize (10)
    xlsChart.easy_getChartDataTable().setHorizontalLines (False)
    xlsChart.easy_getChartDataTable().setLegendKey (True)
    xlsChart.easy_getChartDataTable().getLineColorFormat().setLineColor (CLng(Color.COLOR_BLUE))
    xlsChart.easy_getChartDataTable().setVerticalLines (False)

    ' Export Excel file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial22 - Excel chart datatable.xlsx"
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial22 - Excel chart datatable.xlsx")
    
    ' Confirm export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub

