'---------------------------------------------------------------------
' Tutorial 03
'
' This tutorial shows how to create an Excel file that has
' multiple sheets in VB.NET. The created Excel file is empty and the
' next tutorial shows how to add data into sheets.
'---------------------------------------------------------------------

Imports EasyXLS

Module Tutorial03

    Sub Main()

        Console.WriteLine("Tutorial 03" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that creates Excel files, having two sheets
        Dim workbook As New ExcelDocument(2)

        ' Set the sheet names
        workbook.easy_getSheetAt(0).setSheetName("First tab")
        workbook.easy_getSheetAt(1).setSheetName("Second tab")

        ' Create the Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial03 - create Excel file.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial03 - create Excel file.xlsx")

        ' Confirm the creation of Excel file
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
