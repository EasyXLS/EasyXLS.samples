    '==================================================================
    ' Tutorial 01
    '
    ' This tutorial shows how to export list to Excel file in VBScript.
	' The list contains data from a SQL database.
	' The cells are formatted using a predefined format.
    '==================================================================
    
    ' Constants declaration
    Dim AUTOFORMAT_EASYXLS1
    AUTOFORMAT_EASYXLS1 = 43


    WScript.StdOut.WriteLine("Tutorial 01" & vbcrlf & "----------" & vbcrlf)
    

	' Create an instance of the object that exports Excel files
	Set workbook = CreateObject("EasyXLS.ExcelDocument")

	' Create the database connection
	Dim objConn
	Set objConn = CreateObject("ADODB.Connection")
	objConn.ConnectionString = "Provider=SQLOLEDB;Server=(local);Database=northwind;User ID=sa;Password=;"
	objConn.Open

	' Query the database
	Dim sQueryString
	sQueryString = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', P.ProductName AS 'Product Name', O.UnitPrice AS Price, cast(O.Quantity AS varchar) AS Quantity , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID"
	
	Dim objRS
	Set objRS = CreateObject("ADODB.Recordset") 
	objRS.Open sQueryString, objConn 
	
	' Create the list that stores the query values
	Dim lstRows 
	Set lstRows = CreateObject("EasyXLS.Util.List")
	
	' Add the header row to the list
	Dim	 lstHeaderRow 	
	Set lstHeaderRow  = CreateObject("EasyXLS.Util.List")
	lstHeaderRow.addElement("Order Date")
	lstHeaderRow.addElement("Product Name")
	lstHeaderRow.addElement("Price")
	lstHeaderRow.addElement("Quantity")
	lstHeaderRow.addElement("Value")	
	lstRows.addElement(lstHeaderRow)
	
	' Add the query values from the database to the list
	Do Until objRS.EOF = True
		Set RowList = CreateObject("EasyXLS.Util.List")
		RowList.addElement("" & objRS("Order Date"))
		RowList.addElement("" & objRS("Product Name"))	
		RowList.addElement("" & objRS("Price"))
		RowList.addElement("" & objRS("Quantity"))
		RowList.addElement("" & objRS("Value"))
		lstRows.addElement(RowList)
	    objRS.MoveNext
	Loop 
	

	' Create an instance of the class used to format the cells
	Dim xlsAutoFormat 
	Set xlsAutoFormat = CreateObject("EasyXLS.ExcelAutoFormat")
	xlsAutoFormat.InitAs(AUTOFORMAT_EASYXLS1)
	
	' Export list to Excel file
	WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial01 - export List to Excel.xlsx.")	
	workbook.easy_WriteXLSXFile_FromList_2 "c:\Samples\Tutorial01 - export List to Excel.xlsx", lstRows, xlsAutoFormat, "Sheet1"
 
    ' Confirm export of Excel file
    Dim sError
    sError = workbook.easy_getError()
    If sError = "" Then
		WScript.StdOut.Write(vbcrlf & "File successfully created.")
    Else
		WScript.StdOut.Write(vbcrlf & "Error: " & sError)
    End If
    
	' Close database connection
	objRS.Close
	Set objRS = Nothing
	objConn.Close
	Set objConn = Nothing 
	
    ' Dispose memory
	workbook.Dispose  
	
	Wscript.StdOut.Write(vbcrlf & "Press Enter to exit ...")
	WScript.StdIn.ReadLine() 