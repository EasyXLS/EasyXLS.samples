'---------------------------------------------------------------
' Tutorial 35
'
' This tutorial shows how to read values from a sheet
' of an excel file (For this example we use the file generated
' in Tutorial 09).
'---------------------------------------------------------------

Imports EasyXLS
Imports System.IO

Module Tutorial35

    Sub Main()


        Console.WriteLine("Tutorial 35" & vbCrLf & "----------" & vbCrLf)

        ' Create an instance of the object that generates Excel files
        Dim workbook As New ExcelDocument

        ' Read the file
        Console.WriteLine("Reading file C:\Samples\Tutorial09.xlsx." & vbCrLf)
        Try
            Dim ds As DataSet = workbook.easy_ReadXLSXSheet_AsDataSet("C:\Samples\Tutorial09.xlsx", "First tab")

            ' Display the values
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
