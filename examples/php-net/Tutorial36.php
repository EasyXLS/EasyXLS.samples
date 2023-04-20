<?php
	/*===============================================================================
	 | Tutorial 36 
	 |
	 | This tutorial shows how to read an Excel XLSX file in PHP (the
	 | XLSX file generated by Tutorial 04 as base template), modify
	 | some data and save it to another XLSX file (Tutorial36 - read XLSX file.xlsx).
	  =============================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 36<br>";
	echo "----------<br>";

	// Create an instance of the class that reads Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");
	
	// Read XLSX file
	echo "Reading file: C:\\Samples\\Tutorial04.xlsx<br>";
	if ($workbook->easy_LoadXLSXFile("C:\\Samples\\Tutorial04.xlsx"))
	{
		// Get the table of data for the second worksheet
		$xlsSecondTable = $workbook->easy_getSheet("Second tab")->easy_getExcelTable();
		// Write some data to the second sheet
		$xlsSecondTable->easy_getCell_2("A1")->setValue("Data added by Tutorial36");

		for ($column=0; $column<5; $column++)
		{
			$xlsSecondTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		// Export the new XLSX file
		echo "Writing file: C:\Samples\Tutorial36 - read XLSX file.xlsx<br>";
		$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial36 - read XLSX file.xlsx");
		
		// Confirm export of Excel file
		if ($workbook->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $workbook->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial04.xlsx";
		echo $workbook->easy_getError();
	}
	
	// Dispose memory
	$workbook->Dispose();	
?>
