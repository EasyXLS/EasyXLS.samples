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
    '==============================================================
	' Tutorial 14
	'
	' This tutorial shows how to create an Excel file in VB6 having
	' a sheet and conditional formatting for cell ranges.
	'==============================================================
    
    ConditionalFormatting.Initialize
    DataType.Initialize
    Color.Initialize
    
    Me.Label1.Caption = "Tutorial 14" & vbCrLf & "-----------------" & vbCrLf
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create a sheet
    workbook.easy_addWorksheet_2 ("Sheet1")
    
    ' Get the table of data for the first worksheet
    Set xlsTab = workbook.easy_getSheet("Sheet1")
    Set xlsTable = xlsTab.easy_getExcelTable()

	' Add data in cells
    For i = 0 To 5
        For j = 0 To 3
            If ((i < 2) And (j < 2)) Then
                xlsTable.easy_getCell(i, j).setValue ("12")
            Else
                If ((j = 2) And (i < 2)) Then
                    xlsTable.easy_getCell(i, j).setValue ("1000")
                Else
                    xlsTable.easy_getCell(i, j).setValue ("9")
                End If
            End If
            xlsTable.easy_getCell(i, j).setDataType (DataType.DATATYPE_NUMERIC)
        Next
    Next

    ' Set conditional formatting
    xlsTab.easy_addConditionalFormatting_5 "A1:C3", ConditionalFormatting.CONDITIONALFORMATTING_OPERATOR_BETWEEN, "=9", "=11", True, True, CLng(Color.COLOR_RED)

    ' Set another conditional formatting
    xlsTab.easy_addConditionalFormatting_9 "A6:C6", ConditionalFormatting.CONDITIONALFORMATTING_OPERATOR_BETWEEN, "=COS(PI())+2", "", CLng(Color.COLOR_BISQUE)
    xlsTab.easy_getConditionalFormattingAt_2("A6:C6").getConditionAt(0).setConditionType (ConditionalFormatting.CONDITIONALFORMATTING_CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA)

    ' Export Excel file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial14 - conditional formatting in Excel.xlsx"
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial14 - conditional formatting in Excel.xlsx")
    
    ' Confirm export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub

