'-------------------------------------------------------------
' Tutorial 14
'
' This tutorial shows how to create an Excel file in VB.NET
' having a sheet and conditional formatting for cell ranges.
'-------------------------------------------------------------

Imports System.Drawing
Imports EasyXLS
Imports EasyXLS.Constants

Module Tutorial14

    Sub Main()


        Console.WriteLine("Tutorial 14" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that exports Excel files having one sheet
        Dim workbook As New ExcelDocument(1)

        ' Get the table of data for the first worksheet
        Dim xlsTab As ExcelWorksheet = workbook.easy_getSheet("Sheet1")
        Dim xlsTable = xlsTab.easy_getExcelTable()

        ' Add data in cells
        For i As Integer = 0 To 5
            For j As Integer = 0 To 3
                If ((i < 2) And (j < 2)) Then
                    xlsTable.easy_getCell(i, j).setValue("12")
                Else
                    If ((j = 2) And (i < 2)) Then
                        xlsTable.easy_getCell(i, j).setValue("1000")
                    Else
                        xlsTable.easy_getCell(i, j).setValue("9")
                    End If

                    xlsTable.easy_getCell(i, j).setDataType(DataType.NUMERIC)
                End If
            Next
        Next

        ' Set conditional formatting
        xlsTab.easy_addConditionalFormatting("A1:C3", ConditionalFormatting.OPERATOR_BETWEEN, "=9", "=11", True, True, Color.Red)

        ' Set another conditional formatting
        xlsTab.easy_addConditionalFormatting("A6:C6", ConditionalFormatting.OPERATOR_BETWEEN, "=COS(PI())+2", "", Color.Bisque)
        xlsTab.easy_getConditionalFormattingAt("A6:C6").getConditionAt(0).setConditionType(ConditionalFormatting.CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA)


        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial14 - conditional formatting in Excel.xlsx.")
        workbook.easy_WriteXLSFile("C:\Samples\Tutorial14 - conditional formatting in Excel.xlsx")

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
