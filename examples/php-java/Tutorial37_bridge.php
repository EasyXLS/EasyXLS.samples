<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*================================================================
	 | Tutorial 37
	 |
	 | This tutorial shows how to read an Excel XLS file in PHP (the
	 | XLS file generated by Tutorial 28 as base template), modify
	 | some data and save it to another XLS file (Tutorial37.xls).
	  ==============================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 37<br>";
	echo "----------<br>";
	
	// Create an instance of the class that reads Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Read XLS file
	echo "Reading file: C:\\Samples\\Tutorial28.xls<br>";
	if ($workbook->easy_LoadXLSFile("C:\\Samples\\Tutorial28.xls"))
	{
		// Get the table of data for the second worksheet
		$xlsSecondTable = $workbook->easy_getSheet("Second tab")->easy_getExcelTable();

		// Write some data to the second sheet
		$xlsSecondTable->easy_getCell("A1")->setValue("Data added by Tutorial37");

		for ($column=0; $column<5; $column++)
		{
			$xlsSecondTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		// Export the new XLS file
		echo "Writing file: C:\Samples\Tutorial37 - read XLS file.xls<br>";
		$workbook->easy_WriteXLSFile("C:\Samples\Tutorial37 - read XLS file.xls");
		
		// Confirm export of Excel file
		if ($workbook->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $workbook->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial28.xls";
		echo $workbook->easy_getError();
	}
	
	// Dispose memory
	$workbook->Dispose();	
?>
