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
	' Tutorial 20
	'
	' This tutorial shows how to create an Excel file in VB6
	' and apply an auto-filter to a range of cells.
	'=======================================================
    
    DataType.Initialize
              
    Me.Label1.Caption = "Tutorial 20" & vbCrLf & "-----------------" & vbCrLf
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create a sheet
    workbook.easy_addWorksheet_2 ("Sheet1")
       
    ' Get the table of data for the worksheet
    Set xlsTab = workbook.easy_getSheet("Sheet1")
    Set xlsTable = xlsTab.easy_getExcelTable()
    
    ' Add data in cells for report header
    For Column = 0 To 4
        xlsTable.easy_getCell(0, Column).setValue ("Column " & (Column + 1))
        xlsTable.easy_getCell(0, Column).setDataType (DataType.DATATYPE_STRING)
    Next
    
    ' Add data in cells for report values
    For row = 0 To 99
        For Column = 0 To 4
            xlsTable.easy_getCell(row + 1, Column).setValue ("Data " & (row + 1) & ", " & (Column + 1))
            xlsTable.easy_getCell(row + 1, Column).setDataType (DataType.DATATYPE_STRING)
        Next
    Next
        
    ' Apply auto-filter on cell range A1:E1
    Set xlsFilter = xlsTab.easy_getFilter()
    xlsFilter.setAutoFilter_2 ("A1:E1")

    ' Export Excel file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx"
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx")
    
    ' Confirm export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub

