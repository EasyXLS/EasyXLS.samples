'------------------------------------------------------------
' Tutorial 20
'
' This tutorial shows how to create an Excel file in VB.NET
' and apply an auto-filter to a range of cells.
'------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants


Module Tutorial20

    Sub Main()


        Console.WriteLine("Tutorial 20" & vbCrLf & "----------" & vbCrLf)

        ' Create an instance of the class that exports Excel files having one sheet
        Dim workbook As New ExcelDocument(1)

        ' Set the sheet name
        workbook.easy_getSheetAt(0).setSheetName("Sheet1")

        ' Get the table of data for the worksheet
        Dim xlsTab As ExcelWorksheet = workbook.easy_getSheet("Sheet1")
        Dim xlsTable As ExcelTable = xlsTab.easy_getExcelTable()

        ' Add data in cells for report header
        For column As Integer = 0 To 4
            xlsTable.easy_getCell(0, column).setValue("Column " & (column + 1))
            xlsTable.easy_getCell(0, column).setDataType(DataType.STRING)
        Next

        ' Add data in cells for report values
        For row As Integer = 0 To 99
            For column As Integer = 0 To 4
                xlsTable.easy_getCell(row + 1, column).setValue("Data " & (row + 1) & ", " & (column + 1))
                xlsTable.easy_getCell(row + 1, column).setDataType(DataType.STRING)
            Next
        Next

        ' Apply auto-filter on cell range A1:E1
        Dim xlsFilter As ExcelFilter = xlsTab.easy_getFilter()
        xlsFilter.setAutoFilter("A1:E1")

        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx")

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
