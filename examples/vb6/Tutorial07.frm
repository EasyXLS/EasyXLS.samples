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
      Left            =   240
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
    '============================================================
	' Tutorial 07
	'
	' This code sample shows how to export an Excel file in VB6 
	' having multiple sheets. The first sheet is filled with data 
	' and the cells are formatted and locked.
	' The column header has comments.
	'============================================================
    
    Alignment.Initialize
    Border.Initialize
    DataType.Initialize
    Color.Initialize
    
    Me.Label1.Caption = "Tutorial 07" & vbCrLf & "---------------" & vbCrLf
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create two worksheets
    workbook.easy_addWorksheet_2 ("First tab")
    workbook.easy_addWorksheet_2 ("Second tab")
    
    ' Protect the first sheet
    workbook.easy_getSheetAt(0).setSheetProtected (True)

    ' Get the table of data for the first worksheet
    Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()
    
    ' Create the formatting style for the header
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

    ' Add data in cells for report header
    For Column = 0 To 4
        xlsFirstTable.easy_getCell(0, Column).setValue ("Column " & (Column + 1))
        xlsFirstTable.easy_getCell(0, Column).setStyle (xlsStyleHeader)
                    
        ' Add comment for report header cells
        xlsFirstTable.easy_getCell(0, Column).setComment_2 ("This is column no " & (Column + 1))
    Next
    xlsFirstTable.easy_getRowAt(0).setHeight (30)
    
    ' Create a formatting style for cells
    Set xlsStyleData = CreateObject("EasyXLS.ExcelStyle")
    xlsStyleData.setHorizontalAlignment (Alignment.ALIGNMENT_ALIGNMENT_LEFT)
    xlsStyleData.setForeground (CLng(Color.COLOR_DARKGRAY))
    xlsStyleData.setWrap (False)
    ' Protect cells
	xlsStyleData.setLocked (True)
    xlsStyleData.setDataType (DataType.DATATYPE_STRING)
    
    ' Add data in cells for report values
    For row = 0 To 99
        For Column = 0 To 4
            xlsFirstTable.easy_getCell(row + 1, Column).setValue ("Data " & (row + 1) & ", " & (Column + 1))
            xlsFirstTable.easy_getCell(row + 1, Column).setStyle (xlsStyleData)
        Next
    Next

    ' Set column widths
    xlsFirstTable.setColumnWidth_2 0, 70
    xlsFirstTable.setColumnWidth_2 1, 100
    xlsFirstTable.setColumnWidth_2 2, 70
    xlsFirstTable.setColumnWidth_2 3, 100
    xlsFirstTable.setColumnWidth_2 4, 70

    ' Export Excel file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial07 - cell comment in Excel.xlsx"
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial07 - cell comment in Excel.xlsx")
    
    ' Confirm the export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub

