'-------------------------------------------------------------------------
' Tutorial 34
'
' This tutorial shows how to import Excel to DataSet in VB.NET. The data
' is imported from the active sheet of the Excel file (the Excel file
' generated in Tutorial 09).
'-------------------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants
Imports System.IO

Module Tutorial34

    Sub Main()


        Console.WriteLine("Tutorial 34" & vbCrLf & "----------" & vbCrLf)

        ' Create an instance of the class that imports Excel files
        Dim workbook As New ExcelDocument

        ' Import Excel file to DataSet
        Console.WriteLine("Reading file C:\Samples\Tutorial09.xlsx." & vbCrLf)
        Try
            Dim ds As DataSet = workbook.easy_ReadXLSXActiveSheet_AsDataSet("C:\Samples\Tutorial09.xlsx")

            ' Display imported DataSet values
            Dim dt As DataTable = ds.Tables(0)
            For row As Integer = 0 To dt.Rows.Count - 1
                For column As Integer = 0 To dt.Columns.Count - 1
                    Console.WriteLine("At row " & (row + 1) & ", column " & (column + 1) & _
                     " the value is '" & dt.Rows(row).ItemArray(column) & "'")
                Next
            Next
        Catch ex As Exception
            Console.WriteLine(vbCrLf & "Error reading file C:\Samples\Tutorial09.xlsx " & vbCrLf & workbook.easy_getError())
        End Try

        ' Dispose memory
        workbook.Dispose()

        Console.Write(vbCrLf & "Press Enter to Exit...")
        Console.ReadLine()

    End Sub

End Module
