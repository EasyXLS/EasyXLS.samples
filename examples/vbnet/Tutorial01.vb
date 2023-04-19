'-------------------------------------------------------------------------
' Tutorial 01
'
' This tutorial shows how to export DataSet to Excel file in VB.NET.
' The DataSet contains data from a SQL database, but it also can contain
'data from other sources like GridView, DataGridView, DataGrid or other.
'The cells are formatted using a predefined format.
'-------------------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants
Imports System.IO
Imports System.Data


Module Tutorial01

    Sub Main()


        Console.WriteLine("Tutorial 01" & vbCrLf & "----------" & vbCrLf)

        ' Create an instance of the class that exports Excel files 
        Dim workbook As New ExcelDocument

        ' Create the database connection
        Dim sConnectionString As String = "Initial Catalog=Northwind;Data Source=localhost;User ID=sa;Password=;"
        Dim sqlConnection As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(sConnectionString)
        sqlConnection.Open()

        ' Create the adapter used to fill the dataset
        Dim sQueryString As String = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', "
        sQueryString += " P.ProductName AS 'Product Name', O.UnitPrice AS Price, O.Quantity , O.UnitPrice * O. Quantity AS Value"
        sQueryString += " FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID"
		Dim adp as System.Data.SqlClient.SqlDataAdapter  = new System.Data.SqlClient.SqlDataAdapter(sQueryString, sqlConnection)

        ' Populate the dataset
        Dim ds As DataSet = New DataSet
        adp.Fill(ds)


        ' Export the Excel file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial01 - export DataSet to Excel.xlsx.")
        workbook.easy_WriteXLSXFile_FromDataSet("c:\\Samples\\Tutorial01 - export DataSet to Excel.xlsx", ds, New ExcelAutoFormat(Styles.AUTOFORMAT_EASYXLS1), "Sheet1")

        ' Confirm export of Excel file
        Dim sError As String = workbook.easy_getError()
        If (sError.Equals("")) Then
            Console.Write("File successfully created. Press Enter to Exit...")
        Else
            Console.Write("Error encountered: " + sError + "Press Enter to Exit...")
        End If


        ' Close the database connection
        sqlConnection.Close()

        ' Dispose memory
        workbook.Dispose()
        ds.Dispose()
        sqlConnection.Dispose()
        adp.Dispose()

        Console.ReadLine()

    End Sub

End Module
