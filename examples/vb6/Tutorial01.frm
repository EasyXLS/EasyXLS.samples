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
    '=========================================================
        ' Tutorial 01
        '
        ' This tutorial shows how to export list to Excel file in VB6.
        ' The list contains data from a SQL database.
        ' The cells are formatted using a predefined format.
        '=============================================================
    
    Styles.Initialize

    Me.Label1.Caption = "Tutorial 01" & vbCrLf & "---------------" & vbCrLf
    ' Create an instance of the class that exports Excel files
    Set workbook = CreateObject("EasyXLS.ExcelDocument")
    
    ' Create the database connection
    Dim objConn
    Set objConn = CreateObject("ADODB.Connection")
    objConn.ConnectionString = "Provider=SQLOLEDB;Server=(local);Database=northwind;User ID=sa;Password=;"
    objConn.Open

    ' Query the database
        Dim sQueryString
    sQueryString = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar) + " & _
                   "'/' + CAST(Day(ord.OrderDate) AS varchar) + " & _
                   "'/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', " & _
                   "P.ProductName AS 'Product Name', O.UnitPrice AS Price, " & _
                   "CAST(O.Quantity AS varchar) AS Quantity, " & _
                   "O.UnitPrice * O. Quantity AS Value " & _
                   "FROM Orders AS ord, [Order Details] AS O, Products AS P " & _
                   "WHERE  O.ProductID = P.ProductID AND O.OrderID = ord.OrderID"
    
    Dim objRS
    Set objRS = CreateObject("ADODB.Recordset")
    objRS.Open sQueryString, objConn
    
    ' Create the list that stores the query values
    Dim lstRows
    Set lstRows = CreateObject("EasyXLS.Util.List")
    
    ' Add the report header row to the list
    Dim lstHeaderRow
    Set lstHeaderRow = CreateObject("EasyXLS.Util.List")
    lstHeaderRow.addElement ("Order Date")
    lstHeaderRow.addElement ("Product Name")
    lstHeaderRow.addElement ("Price")
    lstHeaderRow.addElement ("Quantity")
    lstHeaderRow.addElement ("Value")
    lstRows.addElement (lstHeaderRow)
    
    ' Add the query values from the database to the list
    Do Until objRS.EOF = True
        Set RowList = CreateObject("EasyXLS.Util.List")
        RowList.addElement ("" & objRS("Order Date"))
        RowList.addElement ("" & objRS("Product Name"))
        RowList.addElement ("" & objRS("Price"))
        RowList.addElement ("" & objRS("Quantity"))
        RowList.addElement ("" & objRS("Value"))
        lstRows.addElement (RowList)
        objRS.MoveNext
    Loop
    
    ' Create an instance of the class used to format the cells
    Dim xlsAutoFormat
    Set xlsAutoFormat = CreateObject("EasyXLS.ExcelAutoFormat")
    xlsAutoFormat.InitAs (Styles.AUTOFORMAT_EASYXLS1)
    
    ' Export list to Excel file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial01 - export List to Excel.xlsx."
    workbook.easy_WriteXLSXFile_FromList_2 "c:\Samples\Tutorial01 - export List to Excel.xlsx", lstRows, xlsAutoFormat, "Sheet1"

    ' Confirm export of Excel file
    If workbook.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & workbook.easy_getError()
    End If

    ' Close the Recordset object
    objRS.Close
    Set objRS = Nothing
    
    ' Close database connection
    objConn.Close
    Set objConn = Nothing
    
    ' Dispose memory
    workbook.Dispose

End Sub

