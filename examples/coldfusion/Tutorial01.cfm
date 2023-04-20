<!--
==================================================================== 
Tutorial 01

This tutorial shows how to export list to Excel file in Coldfusion.
The list contains data from a SQL database.
The cells are formatted using a predefined format.
====================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="EasyXLS.Constants.Styles" name="Styles" action="CREATE">

	
Tutorial 01<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Query the database -->
<cfquery name="myQuery" datasource="Northwind">
SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order_Date', P.ProductName AS 'Product_Name', O.UnitPrice AS Price, O.Quantity , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID
</cfquery>

<!-- Create the list that stores the query values -->
<cfobject type="java" class="EasyXLS.Util.List" name="lstRows" action="CREATE"> 

<!-- Add the report header row to the list -->
<cfobject type="java" class="EasyXLS.Util.List" name="lstHeaderRow" action="CREATE"> 
<cfset lstHeaderRow.addElement("Order Date")>
<cfset lstHeaderRow.addElement("Product Name")>
<cfset lstHeaderRow.addElement("Price")>
<cfset lstHeaderRow.addElement("Quantity")>
<cfset lstHeaderRow.addElement("Value")>
<cfset lstRows.addElement(lstHeaderRow)>

<!-- Add the query values from the database to the list -->
<cfloop query="myQuery" >
	<cfobject type="java" class="EasyXLS.Util.List" name="RowList" action="CREATE"> 
		<cfset RowList.addElement(#Order_Date#)>
		<cfset RowList.addElement(#Product_Name#)>
		<cfset RowList.addElement(#Price#)>
		<cfset RowList.addElement(#Quantity#)>
		<cfset RowList.addElement(#Value#)>
		<cfset lstRows.addElement(RowList)>
	</cfloop>
	
<!-- Create an instance of the class used to format the cells -->
<cfobject type="java" class="EasyXLS.ExcelAutoFormat" name="xlsAutoFormat" action="CREATE"> 
<cfset xlsAutoFormat.InitAs(Styles.AUTOFORMAT_EASYXLS1)>


<!-- Export list to Excel file -->
Writing file C:\Samples\Tutorial01 - export List to Excel.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile_FromList("C:\Samples\Tutorial01 - export List to Excel.xlsx",lstRows, xlsAutoFormat, "Sheet1")>

<!-- Confirm export of Excel file -->
<cfset sError = workbook.easy_getError()>
<CFIF (sError  IS "")>
  <cfoutput>
	File successfully created.
  </cfoutput>
<CFELSE>
  <cfoutput>
	Error encountered:  #sError#
  </cfoutput>
</CFIF>

<!-- Dispose memory -->
<cfset workbook.Dispose()>
