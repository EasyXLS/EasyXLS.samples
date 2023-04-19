'------------------------------------------------------------------
' Tutorial 19
'
' This tutorial shows how to create an Excel file in VB.NET having
' multiple sheets. The first sheet is filled with data and the
' first cell of the second row contains data in rich text format.
'------------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants


Module Tutorial19

    Sub Main()


        Console.WriteLine("Tutorial 19" & vbCrLf & "----------" & vbCrLf)


        ' Create an instance of the class that exports Excel files having two sheets
        Dim workbook As New ExcelDocument(2)

        ' Set the sheet names
        workbook.easy_getSheetAt(0).setSheetName("First tab")
        workbook.easy_getSheetAt(1).setSheetName("Second tab")

        ' Get the table of data for the first worksheet
        Dim xlsFirstTab As ExcelWorksheet = workbook.easy_getSheetAt(0)
        Dim xlsFirstTable As ExcelTable = xlsFirstTab.easy_getExcelTable()

        ' Create the string used to set the RTF in cell
        Dim sFormattedValue As String
        sFormattedValue = "This is <b>bold</b>."
        sFormattedValue = sFormattedValue & "This is <b>bold</b>."
        sFormattedValue = sFormattedValue & Chr(10) & "This is <i>italic</i>."
        sFormattedValue = sFormattedValue & Chr(10) & "This is <u>underline</u>."
        sFormattedValue = sFormattedValue & Chr(10) & "This is <underline double>double underline</underline double>."
        sFormattedValue = sFormattedValue & Chr(10) & "This is <font color=red>red</font>."
        sFormattedValue = sFormattedValue & Chr(10) & "This is <font color=rgb(255,0,0)>red</font> too."
        sFormattedValue = sFormattedValue & Chr(10) & "This is <font face=""Arial Black"">Arial Black</font>."
        sFormattedValue = sFormattedValue & Chr(10) & "This is <font size=15pt>size 15</font>."
        sFormattedValue = sFormattedValue & Chr(10) & "This is <s>strikethrough</s>."
        sFormattedValue = sFormattedValue & Chr(10) & "This is <sup>superscript</sup>."
        sFormattedValue = sFormattedValue & Chr(10) & "This is <sub>subscript</sub>."
        sFormattedValue = sFormattedValue & Chr(10) & "<b>This</b> <i>is</i> <font color=red face=""Arial Black"" size=15pt><underline double>formatted</underline double></font> <s>text</s>."

        ' Set the rich text value in cell
        xlsFirstTable.easy_getCell(1, 0).setHTMLValue(sFormattedValue)
        xlsFirstTable.easy_getCell(1, 0).setDataType(DataType.STRING)
        xlsFirstTable.easy_getCell(1, 0).setWrap(True)
        xlsFirstTable.easy_getRowAt(1).setHeight(250)
        xlsFirstTable.easy_getColumnAt(0).setWidth(250)

        ' Export Excel file
        Console.WriteLine("Writing file C:\Samples\Tutorial19 - RTF for Excel cells.xlsx.")
        workbook.easy_WriteXLSXFile("C:\Samples\Tutorial19 - RTF for Excel cells.xlsx")

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
