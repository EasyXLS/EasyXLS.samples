<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*===========================================================================

	 | Tutorial 33
	 |
	 | This tutorial shows how to set document properties for Excel file in PHP, 
	 | like 'Subject' property for summary information, 'Manager' property for
	 | document summary information and a custom property.
===============================================================================*/

	include("FileProperty.inc");

	header("Content-Type: text/html");

	echo "Tutorial 33<br>";
	echo "----------<br>";

	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");

	// Create a worksheet
	$workbook->easy_addWorksheet("Sheet1");

	// Set the 'Subject' document property
	$workbook->getSummaryInformation()->setSubject("This is the subject");

	// Set the 'Manager' document property
	$workbook->getDocumentSummaryInformation()->setManager("This is the manager");

	// Set a custom document property
	$workbook->getDocumentSummaryInformation()->setCustomProperty("PropertyName", $VT_NUMBER, "4");

	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial33 - Excel file properties.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial33 - Excel file properties.xlsx");

	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
			echo "File successfully created.";
	else
			echo "Error encountered: " .$workbook->easy_getError();

	// Dispose memory
	$workbook->Dispose();	
?>
