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
	'=============================================================
	' Tutorial 28
	'
	' This tutorial shows how to export data to XLS file that has
	' multiple sheets in VB6. The first sheet is filled with data.
	'=============================================================
    DataType.Initialize
    
    Me.Label1.Caption = "Tutorial 28" & vbCrLf & "---------------" & vbCrLf
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create two sheets
    workbook.easy_addWorksheet_2 ("First tab")
    workbook.easy_addWorksheet_2 ("Second tab")
    
    ' Get the table of data for the first worksheet
    Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()
    
    ' Add data in cells for report header
    For Column = 0 To 4
        xlsFirstTable.easy_getCell(0, Column).setValue ("Column " & (Column + 1))
        xlsFirstTable.easy_getCell(0, Column).setDataType (DataType.DATATYPE_STRING)
    Next
    xlsFirstTable.easy_getRowAt(0).setHeight (30)

    ' Add data in cells for report values
    For row = 0 To 99
        For Column = 0 To 4
            xlsFirstTable.easy_getCell(row + 1, Column).setValue ("Data " & (row + 1) & ", " & (Column + 1))
            xlsFirstTable.easy_getCell(row + 1, Column).setDataType (DataType.DATATYPE_STRING)
        Next
    Next

    ' Set column widths
    xlsFirstTable.setColumnWidth_2 0, 70
    xlsFirstTable.setColumnWidth_2 1, 100
    xlsFirstTable.setColumnWidth_2 2, 70
    xlsFirstTable.setColumnWidth_2 3, 100
    xlsFirstTable.setColumnWidth_2 4, 70
   
    ' Export the XLS file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial28 - export XLS file.xls"
    workbook.easy_WriteXLSFile ("C:\Samples\Tutorial28 - export XLS file.xls")
    
    ' Confirm export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub


