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
	'=====================================================================
	' Tutorial 39
	'
	' This tutorial shows how to convert CSV file to Excel in VB6. The
	' CSV file generated by Tutorial 30 is imported, some data is modified
	' and after that is exported as Excel file.
	'=====================================================================
    
    Me.Label1.Caption = "Tutorial 39" & vbCrLf & "-----------------" & vbCrLf
    
    ' Create an instance of the class used to import/export Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Import CSV file
    Me.Label1.Caption = Me.Label1.Caption & "Reading file: C:\Samples\Tutorial30.csv" & vbCrLf
    If (workbook.easy_LoadCSVFile("C:\Samples\Tutorial30.csv")) Then
        ' Set worksheet name
        workbook.easy_getSheetAt(0).setSheetName ("First tab")

        ' Add new worksheet and add some data in cells (optional step)
        workbook.easy_addWorksheet_2 ("Second tab")
        Set xlsTable = workbook.easy_getSheetAt(1).easy_getExcelTable()
        xlsTable.easy_getCell_2("A1").setValue ("Data added by Tutorial39")

        For Column = 0 To 4
            xlsTable.easy_getCell(1, Column).setValue ("Data " & (Column + 1))
        Next
        
        ' Export Excel file
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial39 - convert CSV to Excel.xlsx"
        workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial39 - convert CSV to Excel.xlsx")
        
        ' Confirm conversion of CSV to Excel
        If workbook.easy_getError() = "" Then
            Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
        Else
            Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
        End If
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error reading file C:\Samples\Tutorial30.csv" & vbCrLf & workbook.easy_getError()
    End If
    
    ' Dispose memory
    workbook.Dispose
End Sub