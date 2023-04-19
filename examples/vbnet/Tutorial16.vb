'-----------------------------------------------------------------------
' Tutorial 16
'
' This tutorial shows how to create an Excel file with image in VB.NET.
' The Excel file has multiple sheets.
' The first worksheet has an image.
'-----------------------------------------------------------------------

Imports EasyXLS

Module Tutorial16

    Sub Main()


        Console.WriteLine("Tutorial 16" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that exports Excel files having two sheets
        Dim workbook As New ExcelDocument(2)

        ' Set the sheet names
        workbook.easy_getSheetAt(0).setSheetName("First tab")
        workbook.easy_getSheetAt(1).setSheetName("Second tab")

        ' Insert image into sheet
        Dim xlsFirstTab As ExcelWorksheet = workbook.easy_getSheetAt(0)
        xlsFirstTab.easy_addImage("C:\\Samples\\EasyXLSLogo.JPG", "A1")

        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial16 - images in Excel.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial16 - images in Excel.xlsx")

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
