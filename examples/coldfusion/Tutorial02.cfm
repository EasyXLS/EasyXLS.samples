<!--
=======================================================================
Tutorial 02

This code sample shows how to export list to Excel file in ColdFusion.
The list contains data from a SQL database.
The cells are formatted using an user-defined format.
=======================================================================
-->

<!-- Constants Classes -->
<cfobject type="java" class="java.awt.Color" name="Color" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Styles" name="Styles" action="CREATE">
<cfobject type="java" class="EasyXLS.Constants.Alignment" name="Alignment" action="CREATE">

	
Tutorial 02<br>
----------<br>


<!-- Create an instance of the class that exports Excel files -->
<cfobject type="java" class="EasyXLS.ExcelDocument" name="workbook" action="CREATE">

<!-- Query the database -->
<cfquery name="myQuery" datasource="Northwind">
SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order_Date', P.ProductName AS 'Product_Name', O.UnitPrice AS Price, ' ' + cast(O.Quantity AS varchar) AS Quantity  , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID
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
	
<!-- Create an instance of the class used to format the cells in the report -->
<cfobject type="java" class="EasyXLS.ExcelAutoFormat" name="xlsAutoFormat" action="CREATE"> 

<!-- Set the formatting style of the header -->
<cfobject type="java" class="EasyXLS.ExcelStyle" name="xlsHeaderStyle" action="CREATE"> 
<cfobject type="java" class="java.awt.Color" name="lightGreen" action="CREATE"> 
<cfset lightGreen.init(JavaCast("int", "144"), JavaCast("int", "238"), JavaCast("int", "144"))>
<cfset xlsHeaderStyle.setBackground(lightGreen)>
<cfset xlsHeaderStyle.setFontSize(12)>
<cfset xlsAutoFormat.setHeaderRowStyle(xlsHeaderStyle)>

<!-- Set the formatting style of the cells (alternating style) -->
<cfobject type="java" class="EasyXLS.ExcelStyle" name="xlsEvenRowStripesStyle" action="CREATE"> 
<cfobject type="java" class="java.awt.Color" name="FloralWhite" action="CREATE"> 
<cfset FloralWhite.init(JavaCast("int", "255"), JavaCast("int", "250"), JavaCast("int", "240"))>
<cfset xlsEvenRowStripesStyle.setBackground(FloralWhite)>
<cfset xlsEvenRowStripesStyle.setFormat("$0.00")>
<cfset xlsEvenRowStripesStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)>
<cfset xlsAutoFormat.setEvenRowStripesStyle(xlsEvenRowStripesStyle)>

<cfobject type="java" class="EasyXLS.ExcelStyle" name="xlsOddRowStripesStyle" action="CREATE"> 
<cfobject type="java" class="java.awt.Color" name="OddRowStripesColor" action="CREATE"> 
<cfset OddRowStripesColor.init(JavaCast("int", "240"), JavaCast("int", "247"), JavaCast("int", "239"))>
<cfset xlsOddRowStripesStyle.setBackground(OddRowStripesColor)>
<cfset xlsOddRowStripesStyle.setFormat("$0.00")>
<cfset xlsOddRowStripesStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)>
<cfset xlsAutoFormat.setOddRowStripesStyle(xlsOddRowStripesStyle)>

<cfobject type="java" class="EasyXLS.ExcelStyle" name="xlsLeftColumnStyle" action="CREATE"> 
<cfobject type="java" class="java.awt.Color" name="FloralWhite" action="CREATE"> 
<cfset FloralWhite.init(JavaCast("int", "255"), JavaCast("int", "250"), JavaCast("int", "240"))>
<cfset xlsLeftColumnStyle.setBackground(FloralWhite)>
<cfset xlsLeftColumnStyle.setFormat("mm/dd/yyyy")>
<cfset xlsLeftColumnStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)>
<cfset xlsAutoFormat.setLeftColumnStyle(xlsLeftColumnStyle)>

<!-- Export list to Excel file -->
Writing file C:\Samples\Tutorial02 - export List to Excel with formatting.xlsx<br>
<cfset ret = workbook.easy_WriteXLSXFile_FromList("C:\Samples\Tutorial02 - export List to Excel with formatting.xlsx",lstRows, xlsAutoFormat, "Sheet1")>

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
