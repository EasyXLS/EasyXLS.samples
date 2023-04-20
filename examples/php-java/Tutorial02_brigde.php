<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*==================================================================
	 | Tutorial 02
	 |
	 | This code sample shows how to export list to Excel file in PHP.
	 | The list contains data from a SQL database.
	 | The cells are formatted using a user-defined format.
	  ================================================================*/
  	
	include("Color.inc");
	include("Alignment.inc");

    // Constants declaration
    $OddRowStripesStyleColor = 0xf0f7ef;
    
	header("Content-Type: text/html");
	
	echo "Tutorial 02<br>";
	echo "----------<br>";
		
	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");

	// Create the database connection
	$serverName = "(local)";
	$connectionInfo = array("Database"=>"northwind","UID"=>"sa","PWD"=>"");
	
	$db_conn = sqlsrv_connect( $serverName, $connectionInfo); 
	if( $db_conn === false )  
	{
   	  echo "Unable to connect.";
  	   die( print_r( sqlsrv_errors(), true));
	}

	// Query the database
	$query_result = sqlsrv_query( $db_conn, "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', P.ProductName AS 'Product Name', O.UnitPrice AS Price, ' ' + cast(O.Quantity AS varchar) AS Quantity , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID" )
		or die( "<strong>ERROR: Query failed</strong>" );

	// Create the list that stores the query values
	$lstRows = new java("EasyXLS.Util.List");
	
	// Add the report header row to the list
	$lstHeaderRow  = new java("EasyXLS.Util.List");
	$lstHeaderRow->addElement("Order Date");
	$lstHeaderRow->addElement("Product Name");
	$lstHeaderRow->addElement("Price");
	$lstHeaderRow->addElement("Quantity");
	$lstHeaderRow->addElement("Value");
	$lstRows->addElement($lstHeaderRow);
			
	// Add the query values from the database to the list
	while ($row=sqlsrv_fetch_array($query_result))
	{
		$RowList = new java("EasyXLS.Util.List");
		$RowList->addElement("" . $row['Order Date']);
		$RowList->addElement("" . $row["Product Name"]);
		$RowList->addElement("" . $row["Price"]);
		$RowList->addElement("" . $row["Quantity"]);
		$RowList->addElement("" . $row["Value"]);
		$lstRows->addElement($RowList);
			
	}
	
	// Create an instance of the class used to format the cells in the report
	$xlsAutoFormat = new java("EasyXLS.ExcelAutoFormat");
	
	// Set the formatting style of the header
	$xlsHeaderStyle = new java("EasyXLS.ExcelStyle");
	$xlsHeaderStyle->setBackground(new java("java.awt.Color", 144, 238, 144));
	$xlsHeaderStyle->setFontSize(12);
	$xlsAutoFormat->setHeaderRowStyle($xlsHeaderStyle);

	// Set the formatting style of the cells (alternating style)
	$xlsEvenRowStripesStyle = new java("EasyXLS.ExcelStyle");
	$xlsEvenRowStripesStyle->setBackground(new java("java.awt.Color", 255, 250, 240));
	$xlsEvenRowStripesStyle->setFormat("$0.00");
	$xlsEvenRowStripesStyle->setHorizontalAlignment($ALIGNMENT_ALIGNMENT_LEFT);
	$xlsAutoFormat->setEvenRowStripesStyle($xlsEvenRowStripesStyle)	;
	$xlsOddRowStripesStyle = new java("EasyXLS.ExcelStyle");
	$xlsOddRowStripesStyle->setBackground(new java("java.awt.Color", 240, 247, 239));
	$xlsOddRowStripesStyle->setFormat("$0.00");
	$xlsOddRowStripesStyle->setHorizontalAlignment ($ALIGNMENT_ALIGNMENT_LEFT);
	$xlsAutoFormat->setOddRowStripesStyle($xlsOddRowStripesStyle);
	$xlsLeftColumnStyle = new java("EasyXLS.ExcelStyle");
	$xlsLeftColumnStyle->setBackground(new java("java.awt.Color", 255, 250, 240));
	$xlsLeftColumnStyle->setFormat("mm/dd/yyyy");
	$xlsLeftColumnStyle->setHorizontalAlignment($ALIGNMENT_ALIGNMENT_LEFT);
	$xlsAutoFormat->setLeftColumnStyle($xlsLeftColumnStyle);
	
	// Export list to Excel file
	echo "Writing file: C:\Samples\Tutorial02 - export List to Excel with formatting.xlsx<br>";
	$workbook->easy_WriteXLSXFile_FromList("C:\Samples\Tutorial02 - export List to Excel with formatting.xlsx", $lstRows, $xlsAutoFormat, "Sheet1");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Free the memory associated with the query
	sqlsrv_free_stmt( $query_result );

	// Close database connection
	sqlsrv_close( $db_conn );           
  
  	// Dispose memory
	$workbook->Dispose();

?>