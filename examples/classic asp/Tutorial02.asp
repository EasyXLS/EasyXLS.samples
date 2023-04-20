<%@ Language=VBScript %>
<!-- #INCLUDE FILE="Color.inc" -->
<!-- #INCLUDE FILE="Alignment.inc" -->

<%
	'=========================================================================
	' Tutorial 02
	'
	' This code sample shows how to export list to Excel file in Classic ASP.
	' The list contains data from a SQL database.
	' The cells are formatted using an user-defined format.
	'=========================================================================

    ' Constants declaration
    Dim OddRowStripesStyleColor    
    OddRowStripesStyleColor = &hfff0f7ef

   
	response.write("Tutorial 02<br>")
	response.write("----------<br>")

	' Create an instance of the class that exports Excel files
	Set workbook = Server.CreateObject("EasyXLS.ExcelDocument")

	' Create the database connection
	DIM objConn
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = "Provider=SQLOLEDB;Server=(local);Database=northwind;User ID=sa;Password=;"
	objConn.Open

	' Query the database
	Dim sQueryString
	sQueryString = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', P.ProductName AS 'Product Name', O.UnitPrice AS Price, ' ' + cast(O.Quantity AS varchar) AS Quantity , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID"
	
	Dim objRS
	Set objRS = Server.CreateObject("ADODB.Recordset") 
	objRS.Open sQueryString, objConn 
	
	' Create the list that stores the query values
	Dim lstRows 
	Set lstRows = CreateObject("EasyXLS.Util.List")
	
	' Add the report header row to the list
	Dim	 lstHeaderRow 	
	Set lstHeaderRow  = Server.CreateObject("EasyXLS.Util.List")
	lstHeaderRow.addElement("Order Date")
	lstHeaderRow.addElement("Product Name")
	lstHeaderRow.addElement("Price")
	lstHeaderRow.addElement("Quantity")
	lstHeaderRow.addElement("Value")	
	lstRows.addElement(lstHeaderRow)
	
	' Add the query values from the database to the list
	Do Until objRS.EOF = True
		set RowList = Server.CreateObject("EasyXLS.Util.List")
		RowList.addElement("" & objRS("Order Date"))
		RowList.addElement("" & objRS("Product Name"))	
		RowList.addElement("" & objRS("Price"))
		RowList.addElement("" & objRS("Quantity"))
		RowList.addElement("" & objRS("Value"))
		lstRows.addElement(RowList)
		objRS.MoveNext
	Loop 
		
	' Create an instance of the class used to format the cells in the report
	Dim xlsAutoFormat 
	set xlsAutoFormat = Server.CreateObject("EasyXLS.ExcelAutoFormat")
	
	' Set the formatting style of the header
	Dim xlsHeaderStyle 
	Set xlsHeaderStyle = Server.CreateObject("EasyXLS.ExcelStyle")
	xlsHeaderStyle.setBackground(CLng(COLOR_LIGHTGREEN))
	xlsHeaderStyle.setFontSize(12)
	xlsAutoFormat.setHeaderRowStyle(xlsHeaderStyle)

	' Set the formatting style of the cells (alternating style)
	Dim xlsEvenRowStripesStyle 
	Set xlsEvenRowStripesStyle = Server.CreateObject("EasyXLS.ExcelStyle")
	xlsEvenRowStripesStyle.setBackground(CLng(COLOR_FLORALWHITE))
	xlsEvenRowStripesStyle.setFormat("$0.00")
	xlsEvenRowStripesStyle.setHorizontalAlignment(ALIGNMENT_ALIGNMENT_LEFT)
	xlsAutoFormat.setEvenRowStripesStyle(xlsEvenRowStripesStyle)	
	Dim xlsOddRowStripesStyle 
	Set xlsOddRowStripesStyle = Server.CreateObject("EasyXLS.ExcelStyle")
	xlsOddRowStripesStyle.setBackground(OddRowStripesStyleColor)
	xlsOddRowStripesStyle.setFormat("$0.00")
	xlsOddRowStripesStyle.setHorizontalAlignment (ALIGNMENT_ALIGNMENT_LEFT)
	xlsAutoFormat.setOddRowStripesStyle(xlsOddRowStripesStyle)
	Dim xlsLeftColumnStyle 
	Set xlsLeftColumnStyle = Server.CreateObject("EasyXLS.ExcelStyle")
	xlsLeftColumnStyle.setBackground(CLng(COLOR_FLORALWHITE))
	xlsLeftColumnStyle.setFormat("mm/dd/yyyy")
	xlsLeftColumnStyle.setHorizontalAlignment (ALIGNMENT_ALIGNMENT_LEFT)
	xlsAutoFormat.setLeftColumnStyle(xlsLeftColumnStyle)
	
	' Export list to Excel file
	response.write("Writing file: C:\Samples\Tutorial02 - export List to Excel with formatting.xlsx<br>")
	workbook.easy_WriteXLSXFile_FromList_2 "C:\Samples\Tutorial02 - export List to Excel with formatting.xlsx", lstRows, xlsAutoFormat, "Sheet1"
	
	' Confirm export of Excel file
	if workbook.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + workbook.easy_getError())
	end if
	
	' Close database connection
	objRS.Close
	Set objRS = Nothing
	objConn.Close
	Set objConn = Nothing 
	
	' Dispose memory
	workbook.Dispose  
%>
