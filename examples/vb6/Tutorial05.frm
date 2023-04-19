VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   2220
   ClientTop       =   1515
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6885
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
    '==========================================================================
    'Tutorial 05
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The first one is full with data
    ' and the cells are formatted.
    '==========================================================================
    
    Alignment.Initialize
    Border.Initialize
    DataType.Initialize
    Color.Initialize
    
    Me.Label1.Caption = "Tutorial 05" & vbCrLf & "---------------" & vbCrLf
    
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Create the worksheets
    xls.easy_addWorksheet_2 ("First tab")
    xls.easy_addWorksheet_2 ("Second tab")
    
    'Get the table of the first worksheet
    Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
    
    'Create the style for the header
    Set xlsStyleHeader = CreateObject("EasyXLS.ExcelStyle")
    xlsStyleHeader.setFont ("Verdana")
    xlsStyleHeader.setFontSize (8)
    xlsStyleHeader.setItalic (True)
    xlsStyleHeader.setBold (True)
    xlsStyleHeader.setForeground (CLng(Color.COLOR_YELLOW))
    xlsStyleHeader.setBackground (CLng(Color.COLOR_BLACK))
    xlsStyleHeader.setBorderColors CLng(Color.COLOR_GRAY), CLng(Color.COLOR_GRAY), CLng(Color.COLOR_GRAY), CLng(Color.COLOR_GRAY)
    xlsStyleHeader.setBorderStyles Border.BORDER_BORDER_MEDIUM, Border.BORDER_BORDER_MEDIUM, Border.BORDER_BORDER_MEDIUM, Border.BORDER_BORDER_MEDIUM
    xlsStyleHeader.setHorizontalAlignment (Alignment.ALIGNMENT_ALIGNMENT_CENTER)
    xlsStyleHeader.setVerticalAlignment (Alignment.ALIGNMENT_ALIGNMENT_BOTTOM)
    xlsStyleHeader.setWrap (True)
    xlsStyleHeader.setDataType (DataType.DATATYPE_STRING)

    'Add the cells for header
    For Column = 0 To 4
        xlsFirstTable.easy_getCell(0, Column).setValue ("Column " & (Column + 1))
        xlsFirstTable.easy_getCell(0, Column).setStyle (xlsStyleHeader)
    Next
    xlsFirstTable.easy_getRowAt(0).setHeight (30)
    
    'Create a style for cells
    Set xlsStyleData = CreateObject("EasyXLS.ExcelStyle")
    xlsStyleData.setHorizontalAlignment (Alignment.ALIGNMENT_ALIGNMENT_LEFT)
    xlsStyleData.setForeground (CLng(Color.COLOR_DARKGRAY))
    xlsStyleData.setWrap (False)
    xlsStyleData.setDataType (DataType.DATATYPE_STRING)
    
    'Add the cells for data
    For row = 0 To 99
        For Column = 0 To 4
            xlsFirstTable.easy_getCell(row + 1, Column).setValue ("Data " & (row + 1) & ", " & (Column + 1))
            xlsFirstTable.easy_getCell(row + 1, Column).setStyle (xlsStyleData)
        Next
    Next

    'Set column widths
    xlsFirstTable.setColumnWidth_2 0, 70
    xlsFirstTable.setColumnWidth_2 1, 100
    xlsFirstTable.setColumnWidth_2 2, 70
    xlsFirstTable.setColumnWidth_2 3, 100
    xlsFirstTable.setColumnWidth_2 4, 70

   
    'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial05 - format Excel cells.xlsx"
    xls.easy_WriteXLSXFile ("C:\Samples\Tutorial05 - format Excel cells.xlsx")
    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    'Dispose memory
    xls.Dispose
End Sub

