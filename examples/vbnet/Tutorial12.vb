'-------------------------------------------------------------------
' Tutorial 12
'
' This tutorial shows how to create an Excel file in VB.NET having
' multiple sheets. The second sheet contains a named area range.
'-------------------------------------------------------------------

Imports EasyXLS

Module Tutorial12

    Sub Main()


        Console.WriteLine("Tutorial 12" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that exports Excel files, having two sheets
        Dim workbook As New ExcelDocument(2)

        ' Set the sheet names
        workbook.easy_getSheetAt(0).setSheetName("First tab")
        workbook.easy_getSheetAt(1).setSheetName("Second tab")

        ' Get the table of data for the second worksheet and populate the worksheet
        Dim xlsSecondTab As ExcelWorksheet = workbook.easy_getSheetAt(1)
        Dim xlsSecondTable = xlsSecondTab.easy_getExcelTable()
        xlsSecondTable.easy_getCell("A1").setValue("Range data 1")
        xlsSecondTable.easy_getCell("A2").setValue("Range data 2")
        xlsSecondTable.easy_getCell("A3").setValue("Range data 3")
        xlsSecondTable.easy_getCell("A4").setValue("Range data 4")

        ' Create a named area range
        xlsSecondTab.easy_addName("Range", "='Second tab'!$A$1:$A$4")


        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial12 - name range in Excel.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial12 - name range in Excel.xlsx")

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
