<%@ Language=VBScript %>

<%
'==========================================================================================
' Tutorial 10
'
' This tutorial shows how to export an Excel file with a merged cell range in ASP classic.
'==========================================================================================

response.Write("Tutorial 10<br>")
response.Write("----------<br>")

' Create an instance of the class that exports Excel files
Set workbook = Server.CreateObject("EasyXLS.ExcelDocument")

' Create a worksheet
workbook.easy_addWorksheet_2("Sheet1")

' Get the table of data for the worksheet
Set xlsTable = workbook.easy_getSheet("Sheet1").easy_getExcelTable()

' Merge cells by range
xlsTable.easy_mergeCells_2("A1:C3")

' Export Excel file
response.Write("Writing file: C:\Samples\Tutorial10 - merge cells in Excel.xlsx<br>")
workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial10 - merge cells in Excel.xlsx")

' Confirm export of Excel file
If workbook.easy_getError() = "" Then
    response.Write("File successfully created.")
Else
    response.Write("Error encountered: " + workbook.easy_getError())
End If

' Dispose memory
workbook.Dispose
%>
