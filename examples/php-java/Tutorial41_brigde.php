<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*==============================================================================
	 | Tutorial 41
	 |
	 | This tutorial shows how to convert XML spreadsheet to Excel in PHP. The
	 | XML Spreadsheet generated by Tutorial 32 is imported, some data is modified
	 | and after that is exported as Excel file.
	  ============================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 41<br>";
	echo "----------<br>";

	// Create an instance of the class used to import/export Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Import XML Spreadsheet file
	echo "Reading file: C:\Samples\Tutorial32.xml<br>";
	if ($workbook->easy_LoadXMLSpreadsheetFile("C:\Samples\Tutorial32.xml"))
	{
		// Get the table of data from the second sheet and add some data in cells (optional step)
		$xlsTable = $workbook->easy_getSheetAt(1)->easy_getExcelTable();
		$xlsTable->easy_getCell("A1")->setValue("Data added by Tutorial41");

		for ($column=0; $column<5; $column++)
		{
			$xlsTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		// Export Excel file
		echo "Writing file: C:\Samples\Tutorial41 - convert XML spreadsheet to Excel.xlsx<br>";
		$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial41 - convert XML spreadsheet to Excel.xlsx");
		
		// Confirm conversion of XML Spreadsheet to Excel
		if ($workbook->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $workbook->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial32_bridge.xml";
		echo $workbook->easy_getError();
	}
	
	// Dispose memory
	$workbook->Dispose();	
?>
