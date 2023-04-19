'---------------------------------------------------------------
' Tutorial 06
'
' This code sample shows how to create an Excel file in VB.NET
' with multiple sheets. The first sheet is protected and
' filled with data. The cells are formatted and locked.
'---------------------------------------------------------------

Imports System.Drawing
Imports EasyXLS
Imports EasyXLS.Constants

Module Tutorial06

    Sub Main()


        Console.WriteLine("Tutorial 06" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that creates Excel files, having two sheets 
        Dim workbook As New ExcelDocument(2)

        ' Set the sheet names
        workbook.easy_getSheetAt(0).setSheetName("First tab")
        workbook.easy_getSheetAt(1).setSheetName("Second tab")

        ' Protect first sheet
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
        ' Protect cells 
        xlsStyleData.setLocked(True)
        xlsFirstTable.easy_setRangeStyle("A2:E101", xlsStyleData)

        ' Set column widths
        xlsFirstTable.setColumnWidth(0, 70)
        xlsFirstTable.setColumnWidth(1, 100)
        xlsFirstTable.setColumnWidth(2, 70)
        xlsFirstTable.setColumnWidth(3, 100)
        xlsFirstTable.setColumnWidth(4, 70)

        ' Create Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial06 - protect Excel sheet.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial06 - protect Excel sheet.xlsx")

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
