<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*=======================================================================
	 | Tutorial 39
	 |
	 | This tutorial shows how to convert CSV file to Excel in PHP. The
	 | CSV file generated by Tutorial 30 is imported, some data is modified
	 | and after that is exported as Excel file.
	  =====================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 39<br>";
	echo "----------<br>";

	// Create an instance of the class used to import/export Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Import CSV file
	echo "Reading file: C:\\Samples\\Tutorial30.csv<br>";
	if ($workbook->easy_LoadCSVFile("C:\\Samples\\Tutorial30.csv"))
	{
		
		// Set worksheet name
		$workbook->easy_getSheetAt(0)->setSheetName("First tab");

		// Add new worksheet and add some data in cells (optional step)
		$workbook->easy_addWorksheet("Second tab");
		$xlsTable = $workbook->easy_getSheetAt(1)->easy_getExcelTable();
		$xlsTable->easy_getCell("A1")->setValue("Data added by Tutorial39");

		for ($column=0; $column<5; $column++)
		{
			$xlsTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		// Export Excel file
		echo "Writing file: C:\Samples\Tutorial39 - convert CSV to Excel.xlsx<br>";
		$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial39 - convert CSV to Excel.xlsx");
		
		// Confirm conversion of CSV to Excel
		if ($workbook->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $workbook->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial30.csv";
		echo $workbook->easy_getError();
	}
	
	// Dispose memory
	$workbook->Dispose();	
?>
