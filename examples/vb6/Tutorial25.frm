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
	'===========================================================
	' Tutorial 25
	'
	' This tutorial shows how to create an Excel file in VB6 and
	' to create a pivot table in a worksheet.
	'===========================================================
	
    DataType.Initialize
    PivotTable.Initialize
    
    Me.Label1.Caption = "Tutorial 25" & vbCrLf & "---------------" & vbCrLf
    
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create two worksheets
    workbook.easy_addWorksheet_2 ("First tab")
    workbook.easy_addWorksheet_2 ("Second tab")
    
    ' Get the table of data for the first worksheet
    Set xlsFirstTable = workbook.easy_getSheetAt(0).easy_getExcelTable()
    
    ' Add data in cells for report header
    xlsFirstTable.easy_getCell(0, 0).setValue ("Sale agent")
    xlsFirstTable.easy_getCell(0, 0).setDataType (DataType.DATATYPE_STRING)
    xlsFirstTable.easy_getCell(0, 1).setValue ("Sale country")
    xlsFirstTable.easy_getCell(0, 1).setDataType (DataType.DATATYPE_STRING)
    xlsFirstTable.easy_getCell(0, 2).setValue ("Month")
    xlsFirstTable.easy_getCell(0, 2).setDataType (DataType.DATATYPE_STRING)
    xlsFirstTable.easy_getCell(0, 3).setValue ("Year")
    xlsFirstTable.easy_getCell(0, 3).setDataType (DataType.DATATYPE_STRING)
    xlsFirstTable.easy_getCell(0, 4).setValue ("Sale amount")
    xlsFirstTable.easy_getCell(0, 4).setDataType (DataType.DATATYPE_STRING)

    xlsFirstTable.easy_getRowAt(0).setBold (True)
    
    ' Add data in cells for report values - the source for pivot table
    xlsFirstTable.easy_getCell(1, 0).setValue ("John Down")
    xlsFirstTable.easy_getCell(1, 1).setValue ("USA")
    xlsFirstTable.easy_getCell(1, 2).setValue ("June")
    xlsFirstTable.easy_getCell(1, 3).setValue ("2010")
    xlsFirstTable.easy_getCell(1, 4).setValue ("550")
        
    xlsFirstTable.easy_getCell(2, 0).setValue ("Scott Valey")
    xlsFirstTable.easy_getCell(2, 1).setValue ("United Kingdom")
    xlsFirstTable.easy_getCell(2, 2).setValue ("June")
    xlsFirstTable.easy_getCell(2, 3).setValue ("2010")
    xlsFirstTable.easy_getCell(2, 4).setValue ("2300")
        
    xlsFirstTable.easy_getCell(3, 0).setValue ("John Down")
    xlsFirstTable.easy_getCell(3, 1).setValue ("USA")
    xlsFirstTable.easy_getCell(3, 2).setValue ("July")
    xlsFirstTable.easy_getCell(3, 3).setValue ("2010")
    xlsFirstTable.easy_getCell(3, 4).setValue ("3100")
        
    xlsFirstTable.easy_getCell(4, 0).setValue ("John Down")
    xlsFirstTable.easy_getCell(4, 1).setValue ("USA")
    xlsFirstTable.easy_getCell(4, 2).setValue ("June")
    xlsFirstTable.easy_getCell(4, 3).setValue ("2011")
    xlsFirstTable.easy_getCell(4, 4).setValue ("1050")
            
    xlsFirstTable.easy_getCell(5, 0).setValue ("John Down")
    xlsFirstTable.easy_getCell(5, 1).setValue ("USA")
    xlsFirstTable.easy_getCell(5, 2).setValue ("July")
    xlsFirstTable.easy_getCell(5, 3).setValue ("2011")
    xlsFirstTable.easy_getCell(5, 4).setValue ("2400")
        
    xlsFirstTable.easy_getCell(6, 0).setValue ("Steve Marlowe")
    xlsFirstTable.easy_getCell(6, 1).setValue ("France")
    xlsFirstTable.easy_getCell(6, 2).setValue ("June")
    xlsFirstTable.easy_getCell(6, 3).setValue ("2011")
    xlsFirstTable.easy_getCell(6, 4).setValue ("1200")
        
    xlsFirstTable.easy_getCell(7, 0).setValue ("Scott Valey")
    xlsFirstTable.easy_getCell(7, 1).setValue ("United Kingdom")
    xlsFirstTable.easy_getCell(7, 2).setValue ("June")
    xlsFirstTable.easy_getCell(7, 3).setValue ("2011")
    xlsFirstTable.easy_getCell(7, 4).setValue ("700")
        
    xlsFirstTable.easy_getCell(8, 0).setValue ("Scott Valey")
    xlsFirstTable.easy_getCell(8, 1).setValue ("United Kingdom")
    xlsFirstTable.easy_getCell(8, 2).setValue ("July")
    xlsFirstTable.easy_getCell(8, 3).setValue ("2011")
    xlsFirstTable.easy_getCell(8, 4).setValue ("360")
        
    ' Create pivot table
    Set xlsPivotTable = CreateObject("EasyXLS.PivotTables.ExcelPivotTable")
            
    xlsPivotTable.setName ("Sales")
    xlsPivotTable.setSourceRange "First tab!$A$1:$E$9", workbook
    xlsPivotTable.setLocation_2 ("A3:G15")
    xlsPivotTable.addFieldToRowLabels ("Sale agent")
    xlsPivotTable.addFieldToColumnLabels ("Year")
    xlsPivotTable.addFieldToValues "Sale amount", "Sale amount per year", PivotTable.PIVOTTABLE_SUBTOTAL_SUM
    xlsPivotTable.addFieldToReportFilter ("Sale country")
    xlsPivotTable.setOutlineForm
    xlsPivotTable.setStyle (PivotTable.PIVOTTABLE_PIVOT_STYLE_DARK_11)
            
    ' Add the pivot table to the second sheet
    Set xlsWorksheet = workbook.easy_getSheet("Second tab")
    xlsWorksheet.easy_addPivotTable (xlsPivotTable)
       
    ' Export Excel file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial25 - pivot table in Excel.xlsx"
    workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial25 - pivot table in Excel.xlsx")
    
    ' Confirm export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Dispose memory
    workbook.Dispose
End Sub


