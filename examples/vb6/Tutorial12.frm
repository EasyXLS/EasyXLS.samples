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
    '===============================================================
	' Tutorial 12
	'
	' This tutorial shows how to create an Excel file in VB6 having
	' multiple sheets. The second sheet contains a named area range.
	'===============================================================
    
    Me.Label1.Caption = "Tutorial 12" & vbCrLf & "-----------------" & vbCrLf
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create two sheets
    workbook.easy_addWorksheet_2 ("First tab")
    workbook.easy_addWorksheet_2 ("Second tab")
    
    ' Get the table of data for the second worksheet and populate the worksheet
    Set xlsSecondTab = workbook.easy_getSheetAt(1)
    Set xlsSecondTable = xlsSecondTab.easy_getExcelTable()
    xlsSecondTable.easy_getCell_2("A1").setValue ("Range data 1")
    xlsSecondTable.easy_getCell_2("A2").setValue ("Range data 2")
    xlsSecondTable.easy_getCell_2("A3").setValue ("Range data 3")
    xlsSecondTable.easy_getCell_2("A4").setValue ("Range data 4")

    ' Create a named area range
    xlsSecondTab.easy_addName_2 "Range", "=Second tab!$A$1:$A$4"
   
    ' Export Excel file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial12 - name range in Excel.xlsx"
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial12 - name range in Excel.xlsx")
    
    ' Confirm export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub



