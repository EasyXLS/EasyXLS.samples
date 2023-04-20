<%@ Language=VBScript %>

<!-- #INCLUDE FILE="FileProperty.inc" -->
<%
'===================================================================================
' Tutorial 33
'
' This tutorial shows how to set document properties for Excel file in ASP classic,
' like 'Subject' property for summary information, 'Manager' property for
' document summary information and a custom property.
'===================================================================================

response.Write("Tutorial 33<br>")
response.Write("----------<br>")

' Create an instance of the class that exports Excel files
Set workbook = Server.CreateObject("EasyXLS.ExcelDocument")

' Create a worksheet
workbook.easy_addWorksheet_2("Sheet1")

' Set the 'Subject' document property
workbook.getSummaryInformation().setSubject("This is the subject")

' Set the 'Manager' document property
workbook.getDocumentSummaryInformation().setManager("This is the manager")

' Set a custom document property
workbook.getDocumentSummaryInformation().setCustomProperty "PropertyName", VT_NUMBER, "4"

' Export Excel file
response.Write("Writing file: C:\Samples\Tutorial33 - Excel file properties.xlsx<br>")
workbook.easy_WriteXLSXFile ("C:\Samples\Tutorial33 - Excel file properties.xlsx")

' Confirm export of Excel file
If workbook.easy_getError() = "" Then
    response.Write("File successfully created.")
Else
    response.Write("Error encountered: " + workbook.easy_getError())
End If

' Dispose memory
workbook.Dispose
%>
