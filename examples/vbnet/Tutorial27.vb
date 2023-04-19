'-------------------------------------------------------------------------------
' Tutorial 27
'
' This tutorial shows how to create an Excel file in VB.NET and
' encrypt the Excel file by setting the password required for opening the file.
'-------------------------------------------------------------------------------

Imports EasyXLS

Module Tutorial27

    Sub Main()


        Console.WriteLine("Tutorial 27" & vbCrLf & "----------" & vbCrLf)

        ' Create an instance of the class that exports Excel files, having two sheets
        Dim workbook As New ExcelDocument(2)

        ' Set the sheet names
        workbook.easy_getSheetAt(0).setSheetName("First tab")
        workbook.easy_getSheetAt(1).setSheetName("Second tab")

        ' Set the password for protecting the Excel file when the file is open
        workbook.easy_getOptions().setPasswordToOpen("password")

        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial27 - protect Excel with password and encryption.xlsx.")
        workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial27 - protect Excel with password and encryption.xlsx")

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
