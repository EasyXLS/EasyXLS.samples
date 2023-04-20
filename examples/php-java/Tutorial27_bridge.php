<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*================================================================================
	 | Tutorial 27
	 |
	 | This tutorial shows how to create an Excel file in PHP and
	 | encrypt the Excel file by setting the password required for opening the file.
	  ==============================================================================*/

	header("Content-Type: text/html");
	
	echo "Tutorial 27<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Create two worksheets
	$workbook->easy_addWorksheet("First tab");
	$workbook->easy_addWorksheet("Second tab");

	// Set the password for protecting the Excel file when the file is open
	$workbook->easy_getOptions()->setPasswordToOpen("password");
		
	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial27 - protect Excel with password and encryption.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial27 - protect Excel with password and encryption.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>
