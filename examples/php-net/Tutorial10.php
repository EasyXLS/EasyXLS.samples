<?php
	/*==================================================================================
	 | Tutorial 10 
	 |
	 | This tutorial shows how to export an Excel file with a merged cell range in PHP.
	  ================================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 10<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");
		
	// Create a worksheet
	$workbook->easy_addWorksheet_2("Sheet1");

	// Get the table of data for the worksheet
	$xlsTable = $workbook->easy_getSheet("Sheet1")->easy_getExcelTable();

	// Merge cells by range
	$xlsTable->easy_mergeCells_2("A1:C3");
	
	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial10 - merge cells in Excel.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial10 - merge cells in Excel.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>