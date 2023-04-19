'----------------------------------------------------------------------------------
' Tutorial 17
'
' This tutorial shows how to create an Excel file with groups on rows in VB.NET.
' The Excel file has two worksheets. The first one is full with data and contains
' the data groups.
'----------------------------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants


Module Tutorial17

    Sub Main()


        Console.WriteLine("Tutorial 17" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that exports Excel files having two sheets
        Dim workbook As New ExcelDocument(2)

        ' Set the sheet names
        workbook.easy_getSheetAt(0).setSheetName("First tab")
        workbook.easy_getSheetAt(1).setSheetName("Second tab")

        ' Get the table of data for the first worksheet
        Dim xlsFirstTab As ExcelWorksheet = workbook.easy_getSheetAt(0)
        Dim xlsFirstTable = xlsFirstTab.easy_getExcelTable()

        ' Add data in cells for report header
        For column As Integer = 0 To 4
            xlsFirstTable.easy_getCell(0, column).setValue("Column " & (column + 1))
            xlsFirstTable.easy_getCell(0, column).setDataType(DataType.STRING)
        Next
        xlsFirstTable.easy_getRowAt(0).setHeight(30)

        ' Add data in cells for report values
        For row As Integer = 0 To 24
            For column As Integer = 0 To 4
                xlsFirstTable.easy_getCell(row + 1, column).setValue("Data " & (row + 1) & ", " & (column + 1))
                xlsFirstTable.easy_getCell(row + 1, column).setDataType(DataType.STRING)
            Next
        Next

        ' Set column widths
        xlsFirstTable.setColumnWidth(0, 70)
        xlsFirstTable.setColumnWidth(1, 100)
        xlsFirstTable.setColumnWidth(2, 70)
        xlsFirstTable.setColumnWidth(3, 100)
        xlsFirstTable.setColumnWidth(4, 70)

        ' Group rows and format A1:E26 cell range
        Dim xlsFirstDataGroup As New ExcelDataGroup("A1:E26", DataGroup.GROUP_BY_ROWS, False)
        xlsFirstDataGroup.setAutoFormat(New ExcelAutoFormat(EasyXLS.Constants.Styles.AUTOFORMAT_EASYXLS1))
        xlsFirstTab.easy_addDataGroup(xlsFirstDataGroup)

        ' Group rows and format A2:E10 cell range, outline level two, inside previous group
        Dim xlsSecondDataGroup As New ExcelDataGroup("A2:E10", DataGroup.GROUP_BY_ROWS, False)
        xlsSecondDataGroup.setAutoFormat(New ExcelAutoFormat(EasyXLS.Constants.Styles.AUTOFORMAT_EASYXLS2))
        xlsFirstTab.easy_addDataGroup(xlsSecondDataGroup)

        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial17 - group data in Excel.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial17 - group data in Excel.xlsx")

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
