'---------------------------------------------------------------
' Tutorial 07
'
' This code sample shows how to export an Excel file in VB.NET
' having multiple sheets. The first sheet is filled with data
' and the cells are formatted and locked.
' The column header has comments.
'---------------------------------------------------------------

Imports System.Drawing
Imports EasyXLS
Imports EasyXLS.Constants

Module Tutorial07

    Sub Main()


        Console.WriteLine("Tutorial 07" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that exports Excel files, having two sheets
        Dim workbook As New ExcelDocument(2)

        ' Set the sheet names
        workbook.easy_getSheetAt(0).setSheetName("First tab")
        workbook.easy_getSheetAt(1).setSheetName("Second tab")

        ' Protect the first sheet
        workbook.easy_getSheetAt(0).setSheetProtected(True)

        ' Get the table of data for the first worksheet
        Dim xlsFirstTab As ExcelWorksheet = workbook.easy_getSheetAt(0)
        Dim xlsFirstTable = xlsFirstTab.easy_getExcelTable()

        ' Create the formatting style for the header
        Dim xlsStyleHeader As New ExcelStyle("Verdana", 8, True, True, Color.Yellow)
        xlsStyleHeader.setBackground(Color.Black)
        xlsStyleHeader.setBorderColors(Color.Gray, Color.Gray, Color.Gray, Color.Gray)
        xlsStyleHeader.setBorderStyles(Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM)
        xlsStyleHeader.setHorizontalAlignment(Alignment.ALIGNMENT_CENTER)
        xlsStyleHeader.setVerticalAlignment(Alignment.ALIGNMENT_BOTTOM)
        xlsStyleHeader.setWrap(True)
        xlsStyleHeader.setDataType(DataType.STRING)

        ' Add data in cells for report header
        For column As Integer = 0 To 4
            xlsFirstTable.easy_getCell(0, column).setValue("Column " & (column + 1))
            xlsFirstTable.easy_getCell(0, column).setStyle(xlsStyleHeader)

            ' Add comment for report header cells 
            xlsFirstTable.easy_getCell(0, column).setComment("This is column no " & (column + 1))
        Next
        xlsFirstTable.easy_getRowAt(0).setHeight(30)


        ' Add data in cells for report values
        For row As Integer = 0 To 99
            For column As Integer = 0 To 4
                xlsFirstTable.easy_getCell(row + 1, column).setValue("Data " & (row + 1) & ", " & (column + 1))
            Next
        Next

        ' Create a formatting style for cells 
        Dim xlsStyleData As New ExcelStyle
        xlsStyleData.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)
        xlsStyleData.setForeground(Color.DarkGray)
        xlsStyleData.setWrap(False)
        xlsStyleData.setDataType(DataType.STRING)
        ' Protect(cells)
        xlsStyleData.setLocked(True)
        xlsFirstTable.easy_setRangeStyle("A2:E101", xlsStyleData)

        ' Set column widths
        xlsFirstTable.setColumnWidth(0, 70)
        xlsFirstTable.setColumnWidth(1, 100)
        xlsFirstTable.setColumnWidth(2, 70)
        xlsFirstTable.setColumnWidth(3, 100)
        xlsFirstTable.setColumnWidth(4, 70)

        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial07 - cell comment in Excel.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial07 - cell comment in Excel.xlsx")

        ' Confirm the export of Excel file
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