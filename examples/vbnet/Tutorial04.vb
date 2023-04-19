'------------------------------------------------------------------
' Tutorial 04
'
' This tutorial shows how to export data to XLSX file that has
' multiple sheets in VB.NET. The first sheet is filled with data.
'------------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants

Module Tutorial04

    Sub Main()


        Console.WriteLine("Tutorial 04" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that exports Excel files, having two sheets
        Dim workbook As New ExcelDocument(2)

        ' Set the sheet names
        workbook.easy_getSheetAt(0).setSheetName("First tab")
        workbook.easy_getSheetAt(1).setSheetName("Second tab")

        ' Get the table of data for the first worksheet
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

        ' Set column widths
        xlsFirstTable.setColumnWidth(0, 70)
        xlsFirstTable.setColumnWidth(1, 100)
        xlsFirstTable.setColumnWidth(2, 70)
        xlsFirstTable.setColumnWidth(3, 100)
        xlsFirstTable.setColumnWidth(4, 70)

        ' Export the XLSX file
        Console.WriteLine("Writing file C:\Samples\Tutorial04 - export data to Excel.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial04 - export data to Excel.xlsx")

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
