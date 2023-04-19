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
    '=================================================================================
	' Tutorial 10
	'
	' This tutorial shows how to export an Excel file with a merged cell range in VB6.
	'=================================================================================
    
    Me.Label1.Caption = "Tutorial 10" & vbCrLf & "---------------" & vbCrLf
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create a worksheet
    workbook.easy_addWorksheet_2 ("Sheet1")
    
    ' Get the table of data for the worksheet
    Set xlsTable = workbook.easy_getSheet("Sheet1").easy_getExcelTable()
    
    ' Merge cells by range
    xlsTable.easy_mergeCells_2 ("A1:C3")

    ' Export Excel file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial10 - merge cells in Excel.xlsx"
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial10 - merge cells in Excel.xlsx")
    
    ' Confirm export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub

