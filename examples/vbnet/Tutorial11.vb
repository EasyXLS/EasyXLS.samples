'----------------------------------------------------------------
' Tutorial 11
'
' This tutorial shows how to create an Excel file in VB.NET that
' has a cell that contains SUM formula for a range of cells.
'----------------------------------------------------------------

Imports EasyXLS

Module Tutorial11

    Sub Main()


        Console.WriteLine("Tutorial 11" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that exports Excel files
        Dim workbook As New ExcelDocument

        ' Create a sheet
        workbook.easy_addWorksheet("Formula")

        ' Get the table of data for the sheet, add data in sheet and the formula
        Dim xlsFirstTab As ExcelWorksheet = workbook.easy_getSheet("Formula")
        Dim xlsTable = xlsFirstTab.easy_getExcelTable()
        xlsTable.easy_getCell("A1").setValue("1")
        xlsTable.easy_getCell("A2").setValue("2")
        xlsTable.easy_getCell("A3").setValue("3")
        xlsTable.easy_getCell("A4").setValue("4")
        xlsTable.easy_getCell("A6").setValue("=SUM(A1:A4)")

        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial11 - formulas in Excel.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial11 - formulas in Excel.xlsx")

        ' Confirm export of Excel file
        Dim sError As String = workbook.easy_getError()
        If (sError.Equals("")) Then
            Console.Write(vbCrLf & "File successfully created. Press Enter to Exit...")
        Else
            Console.Write(vbCrLf & "Error encountered: " & sError & vbCrLf & "Press Enter to Exit...")
        End If
        Console.ReadLine()

    End Sub

End Module
