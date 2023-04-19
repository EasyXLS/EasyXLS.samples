'------------------------------------------------------------------------------
' Tutorial 33
'
' This tutorial shows how to set document properties for Excel file in VB.NET,
' like 'Subject' property for summary information, 'Manager' property for
' document summary information and a custom property.
'------------------------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants


Module Tutorial33

    Sub Main()

        Console.WriteLine("Tutorial 33" & vbCrLf & "----------" & vbCrLf)

        ' Create an instance of the class that exports Excel files
        Dim workbook As New ExcelDocument(1)

        ' Set the 'Subject' document property
        workbook.getSummaryInformation().setSubject("This is the subject")

        ' Set the 'Manager' document property
        workbook.getDocumentSummaryInformation().setManager("This is the manager")

        ' Set a custom document property
        workbook.getDocumentSummaryInformation().setCustomProperty("PropertyName", FileProperty.VT_NUMBER, "4")

        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial33 - Excel file properties.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial33 - Excel file properties.xlsx")

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
