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
	'===============================================================================
	' Tutorial 36
	'
	' This tutorial shows how to read an Excel XLSX file in VB6
	' (the XLSX file generated by Tutorial 04 as base template), modify
	' some data and save it to another XLSX file (Tutorial36 - read XLSX file.xlsx).
	'===============================================================================
    
    Me.Label1.Caption = "Tutorial 36" & vbCrLf & "-----------------" & vbCrLf
    
    ' Create an instance of the class that reads Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Read XLSX file
    Me.Label1.Caption = Me.Label1.Caption & "Reading file: C:\Samples\Tutorial04.xlsx" & vbCrLf
    If (workbook.easy_LoadXLSXFile("C:\Samples\Tutorial04.xlsx")) Then
                
        ' Get the table of data for the second worksheet
        Set xlsSecondTable = workbook.easy_getSheetAt(1).easy_getExcelTable()
        
        ' Write some data to the second sheet
        xlsSecondTable.easy_getCell_2("A1").setValue ("Data added by Tutorial36")

        For Column = 0 To 4
            xlsSecondTable.easy_getCell(1, Column).setValue ("Data " & (Column + 1))
        Next
        
        ' Export the new XLSX file
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial36 - read XLSX file.xlsx"
        workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial36 - read XLSX file.xlsx")
        
        ' Confirm export of Excel file
        If workbook.easy_getError() = "" Then
            Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
        Else
            Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
        End If
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error reading file C:\Samples\Tutorial04.xlsx" & vbCrLf & workbook.easy_getError()
    End If
    
    ' Dispose memory
    workbook.Dispose
End Sub