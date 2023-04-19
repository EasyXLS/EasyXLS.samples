'----------------------------------------------------------------
' Tutorial 30
'
' This tutorial shows how to export data to CSV file in VB.NET.
'----------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants

Module Tutorial30

    Sub Main()


        Console.WriteLine("Tutorial 30" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that exports Excel files, having a sheet
        Dim workbook As New ExcelDocument(2)

        ' Set the sheet name
        workbook.easy_getSheetAt(0).setSheetName("First tab")

        ' Get the table of data for the worksheet
        Dim xlsFirstTab As ExcelWorksheet = workbook.easy_getSheetAt(0)
        Dim xlsFirstTable = xlsFirstTab.easy_getExcelTable()

        ' Add data in cells for report header
        For column As Integer = 0 To 4
            xlsFirstTable.easy_getCell(0, column).setValue("Column " & (column + 1))
            xlsFirstTable.easy_getCell(0, column).setDataType(DataType.STRING)
        Next

        ' Add data in cells for report values
        For row As Integer = 0 To 99
            For column As Integer = 0 To 4
                xlsFirstTable.easy_getCell(row + 1, column).setValue("Data " & (row + 1) & ", " & (column + 1))
                xlsFirstTable.easy_getCell(row + 1, column).setDataType(DataType.STRING)
            Next
        Next

        ' Export CSV file
        Console.WriteLine("Writing file C:\Samples\Tutorial30 - export CSV file.csv.")
        workbook.easy_WriteCSVFile("C:\Samples\Tutorial30 - export CSV file.csv", "First tab")

        ' Confirm export of CSV file
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
