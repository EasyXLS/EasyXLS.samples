'-------------------------------------------------------------------------
' Tutorial 02
'
' This code sample shows how to export DataSet to Excel file in VB.NET.
' The DataSet contains data from a SQL database, but it also can contain
' data from other sources like GridView, DataGridView, DataGrid or other.
' The cells are formatted using a user-defined format.
'-------------------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants
Imports System.IO
Imports System.Data
Imports System.Drawing



Module Tutorial02

    Sub Main()


        Console.WriteLine("Tutorial 02" & vbCrLf & "----------" & vbCrLf)

        ' Create an instance of the class that exports Excel files 
        Dim workbook As New ExcelDocument

        ' Create the database connection
        Dim sConnectionString As String = "Initial Catalog=Northwind;Data Source=localhost;User ID=sa;Password=;"
        Dim sqlConnection As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(sConnectionString)
        sqlConnection.Open()

        ' Create the adapter used to fill the dataset
        Dim sQueryString As String = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', "
        sQueryString += " P.ProductName AS 'Product Name', O.UnitPrice AS Price, cast(O.Quantity AS varchar) As Quantity , O.UnitPrice * O. Quantity AS Value"
        sQueryString += " FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID"
        Dim adp As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(sQueryString, sqlConnection)

        ' Populate the dataset
        Dim dataset As DataSet = New DataSet
        adp.Fill(dataset)

        ' Create an instance of the class used to format the cells in the report
        Dim xlsAutoFormat As ExcelAutoFormat = New ExcelAutoFormat

        ' Set the formatting style of the cells (alternating style)
        Dim xlsHeaderStyle As ExcelStyle = New ExcelStyle(Color.LightGreen)
        xlsHeaderStyle.setFontSize(12)
        xlsAutoFormat.setHeaderRowStyle(xlsHeaderStyle)
        Dim xlsEvenRowStripesStyle As ExcelStyle = New ExcelStyle(Color.FloralWhite)
        xlsEvenRowStripesStyle.setFormat("$0.00")
        xlsEvenRowStripesStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)
        xlsAutoFormat.setEvenRowStripesStyle(xlsEvenRowStripesStyle)
        Dim xlsOddRowStripesStyle As ExcelStyle = New ExcelStyle(Color.FromArgb(240, 247, 239))
        xlsOddRowStripesStyle.setFormat("$0.00")
        xlsOddRowStripesStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)
        xlsAutoFormat.setOddRowStripesStyle(xlsOddRowStripesStyle)
        Dim xlsLeftColumnStyle As ExcelStyle = New ExcelStyle(Color.FloralWhite)
        xlsLeftColumnStyle.setFormat("mm/dd/yyyy")
        xlsLeftColumnStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)
        xlsAutoFormat.setLeftColumnStyle(xlsLeftColumnStyle)

        ' Export the Excel file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial02 - export DataSet to Excel with formatting.xlsx.")
        workbook.easy_WriteXLSXFile_FromDataSet("c:\\Samples\\Tutorial02 - export DataSet to Excel with formatting.xlsx", dataset, xlsAutoFormat, "Sheet1")

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
        dataset.Dispose()
        sqlConnection.Dispose()
        adp.Dispose()

        Console.ReadLine()

    End Sub

End Module
