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
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   0
      Width           =   6855
   End
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
    '===================================================================
	' Tutorial 35
	'
	' This tutorial shows how to import Excel sheet to List in VB6.
	' The data is imported from a specific Excel sheet (For this example
	' we use the Excel file generated in Tutorial 09).
	'===================================================================
    
    Me.Text1 = "Tutorial 35" & vbCrLf & "-----------------" & vbCrLf
    
    ' Create an instance of the class that imports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Import Excel sheet to List
    Me.Text1 = Me.Text1 & "Reading file: C:\Samples\Tutorial09.xlsx" & vbCrLf
    Set rows = workbook.easy_ReadXLSXSheet_AsList_3("C:\Samples\Tutorial09.xlsx", "First tab")

	' Confirm import of Excel file
    If workbook.easy_getError() = "" Then
        ' Display imported List values
        For rowIndex = 0 To rows.Size() - 1
            Set row = rows.elementAt(rowIndex)
            For cellIndex = 0 To row.Size - 1
                Me.Text1 = Me.Text1 & "At row " & (rowIndex + 1) & ", column " & (cellIndex + 1) & " the value is '" & row.elementAt(cellIndex) & "'" & vbCrLf
            Next
        Next
    Else
        Me.Text1 = Me.Text1 & vbCrLf & "Error reading file C:\Samples\Tutorial09.xls " & vbCrLf & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub


