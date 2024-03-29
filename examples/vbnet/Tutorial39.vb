'-----------------------------------------------------------------------
' Tutorial 39
'
' This tutorial shows how to convert CSV file to Excel in VB.NET. The
' CSV file generated by Tutorial 30 is imported, some data is modified
' and after that is exported as Excel file.
'-----------------------------------------------------------------------

Imports EasyXLS
Imports System.IO

Module Tutorial39

    Sub Main()


        Console.WriteLine("Tutorial 39" & vbCrLf & "----------" & vbCrLf)

        ' Create an instance of the class used to import/export Excel files
        Dim workbook As New ExcelDocument

        ' Import CSV file
        Console.WriteLine("Reading file C:\Samples\Tutorial30.csv." & vbCrLf)
        If (workbook.easy_LoadCSVFile("C:\Samples\Tutorial30.csv")) Then

            ' Set worksheet name
            workbook.easy_getSheetAt(0).setSheetName("First tab")

            ' Add new worksheet and add some data in cells (optional step)
            workbook.easy_addWorksheet("Second tab")
            Dim xlsSecondTab As ExcelWorksheet = workbook.easy_getSheetAt(1)
            Dim xlsTable = xlsSecondTab.easy_getExcelTable

            xlsTable.easy_getCell("A1").setValue("Data added by Tutorial39")

            For column As Integer = 0 To 4
                xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
            Next

            ' Export Excel file
            Console.WriteLine(vbCrLf & "Writing file C:\Samples\Tutorial39 - convert CSV to Excel.xlsx.")
            workbook.easy_WriteXLSXFile("C:\Samples\Tutorial39 - convert CSV to Excel.xlsx")

            ' Confirm conversion of CSV to Excel
            Dim sError As String = workbook.easy_getError()
            If (sError.Equals("")) Then
                Console.Write(vbCrLf & "File successfully created.")
            Else
                Console.Write(vbCrLf & "Error encountered: " & sError)
            End If
        Else
            Console.WriteLine(vbCrLf & "Error reading file C:\Samples\Tutorial30.csv " & vbCrLf & workbook.easy_getError())
        End If

        ' Dispose memory
        workbook.Dispose()

        Console.WriteLine(vbCrLf & "Press Enter to Exit...")
        Console.ReadLine()
    End Sub

End Module
