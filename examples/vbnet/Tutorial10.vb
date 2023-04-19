'-------------------------------------------------------------------------------------
' Tutorial 10
'
' This tutorial shows how to export an Excel file with a merged cell range in VB.NET.
'-------------------------------------------------------------------------------------

Imports EasyXLS

Module Tutorial10

    Sub Main()


        Console.WriteLine("Tutorial 10" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that exports Excel files
        Dim workbook As New ExcelDocument(1)

        ' Get the table of data for the worksheet
        Dim xlsFirstTab As ExcelWorksheet = workbook.easy_getSheet("Sheet1")
        Dim xlsTable = xlsFirstTab.easy_getExcelTable()

        ' Merge cells by range
        xlsTable.easy_mergeCells("A1:C3")

        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial10 - merge cells in Excel.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial10 - merge cells in Excel.xlsx")

        ' Confirm export of Excel file
        Dim sError As String = workbook.easy_getError()
        If (sError.Equals("")) Then
            Console.Write(vbCrLf & "File successfully created. Press Enter to Exit...")
        Else
            Console.Write(vbCrLf & "Error encountered: " & sError & vbCrLf & "Press Enter to Exit...")
        End If

        ' Dispose memory
        workbook.Dispose()

        Console.ReadLine()

    End Sub

End Module
