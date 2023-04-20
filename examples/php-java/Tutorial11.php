<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*==============================================================
	 | Tutorial 11
	 |
	 | This tutorial shows how to create an Excel file in PHP that
	 | has a cell that contains SUM formula for a range of cells.
	  ============================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 11<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Create a sheet
	$workbook->easy_addWorksheet("Formula");
		
	// Get the table of data for the sheet, add data in sheet and the formula
	$xlsTable = $workbook->easy_getSheet("Formula")->easy_getExcelTable();
	$xlsTable->easy_getCell("A1")->setValue("1");
	$xlsTable->easy_getCell("A2")->setValue("2");
	$xlsTable->easy_getCell("A3")->setValue("3");
	$xlsTable->easy_getCell("A4")->setValue("4");
	$xlsTable->easy_getCell("A6")->setValue("=SUM(A1:A4)");
	
	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial11 - formulas in Excel.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial11 - formulas in Excel.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>