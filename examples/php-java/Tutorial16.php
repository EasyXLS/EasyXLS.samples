<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*====================================================================
	 | Tutorial 16
	 |
	 | This tutorial shows how to create an Excel file with image in PHP
	 | The Excel file has multiple sheets.
	 | The first sheet has an image inserted.
	  ==================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 16<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Create two sheets
	$workbook->easy_addWorksheet("First tab");
	$workbook->easy_addWorksheet("Second tab");

	// Insert image into sheet
	$workbook->easy_getSheetAt(0)->easy_addImage("C:\\Samples\\EasyXLSLogo.JPG", "A1");
	
	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial16 - images in Excel.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial16 - images in Excel.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>