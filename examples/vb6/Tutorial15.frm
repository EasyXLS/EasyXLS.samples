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
	'====================================================================================
	' Tutorial 15
	'
	' This tutorial shows how to create an Excel file with hyperlinks in VB6.
	'
	' EasyXLS supports the following hyperlink types:
	'		1 - hyperlink to URL
	'		2 - hyperlink to file
	'		3 - hyperlink to UNC
	'		4 - hyperlink to cell in the same Excel file
	'		5 - hyperlink to name
	'
	' The link can be placed on a range of cells.
	'
	' Every type of hyperlink accepts a tool tip description.
	'
	' Every type of hyperlink accepts a text mark. A text mark is a link inside the file.
	'====================================================================================
    
    HyperlinkType.Initialize
    
    Me.Label1.Caption = "Tutorial 15" & vbCrLf & "-----------------" & vbCrLf
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create two sheets
    workbook.easy_addWorksheet_2 ("First tab")
    workbook.easy_addWorksheet_2 ("Second tab")
    
    Set xlsTab1 = workbook.easy_getSheetAt(0)
    Set xlsTab2 = workbook.easy_getSheetAt(1)
    
    ' Create hyperlink to URL
    xlsTab1.easy_addHyperlink_3 HyperlinkType.HYPERLINKTYPE_URL, "http://www.euoutsourcing.com", "Link to URL", "B2:E2"

    ' Create hyperlink to file
    xlsTab1.easy_addHyperlink_3 HyperlinkType.HYPERLINKTYPE_FILE, "c:\tutorial27.xls", "Link to file", "B3"

    ' Create hyperlink to UNC
    xlsTab1.easy_addHyperlink_3 HyperlinkType.HYPERLINKTYPE_UNC, "\\nicoar\samples\tutorial9.xls", "Link to UNC", "B4:D4"

    ' Create hyperlink to cell on second sheet
    xlsTab1.easy_addHyperlink_3 HyperlinkType.HYPERLINKTYPE_CELL, "'Second tab'!D3", "Link to CELL", "B5"

    ' Create a name on the second sheet
    xlsTab2.easy_addName_2 "Name", "=Second tab!$A$1:$A$4"
    
    ' Create hyperlink to name
    xlsTab1.easy_addHyperlink_3 HyperlinkType.HYPERLINKTYPE_CELL, "Name", "Link to a name", "B6"

    ' Export Excel file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial15 - hyperlinks in Excel.xlsx"
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial15 - hyperlinks in Excel.xlsx")
    
    ' Confirm export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub
