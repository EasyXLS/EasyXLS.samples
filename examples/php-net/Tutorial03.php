<?php
	/*==================================================================
	 | Tutorial 03
	 |
	 | This tutorial shows how to create an Excel file that has 
	 | multiple sheets in PHP. The created Excel file is empty and the 
	 | next tutorial shows how to add data into sheets.
	  ================================================================*/

	header("Content-Type: text/html");
	
	echo "Tutorial 03<br>";
	echo "----------<br>";
	
	// Create an instance of the class that creates Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");
	
	// Create two sheets 
	$workbook->easy_addWorksheet_2("First tab");
	$workbook->easy_addWorksheet_2("Second tab");
	
	// Create Excel file
	echo "Writing file: C:\Samples\Tutorial03 - create Excel file.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial03 - create Excel file.xlsx");
	
	// Confirm the creation of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>