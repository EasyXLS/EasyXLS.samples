<?php
	/*===============================================================================
	 | Tutorial 38
	 |
	 | This tutorial shows how to read an Excel XLSB file in PHP (the
	 | XLSB file generated by Tutorial 29 as base template), modify
	 | some data and save it to another XLSB file (Tutorial38 - read XLSB file.xlsb).
	  =============================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 38<br>";
	echo "----------<br>";

	// Create an instance of the class that reads Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");
	
	// Read XLSB file
	echo "Reading file: C:\\Samples\\Tutorial29.xlsb<br>";
	if ($workbook->easy_LoadXLSBFile("C:\\Samples\\Tutorial29.xlsb"))
	{
		// Get the table of data for the second worksheet
		$xlsSecondTable = $workbook->easy_getSheet("Second tab")->easy_getExcelTable();
		// Write some data to the second sheet
		$xlsSecondTable->easy_getCell_2("A1")->setValue("Data added by Tutorial38");

		for ($column=0; $column<5; $column++)
		{
			$xlsSecondTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		// Export the new XLSB file
		echo "Writing file: C:\Samples\Tutorial38 - read XLSB file.xlsb<br>";
		$workbook->easy_WriteXLSBFile("C:\Samples\Tutorial38 - read XLSB file.xlsb");
		
		// Confirm export of Excel file
		if ($workbook->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $workbook->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial29.xlsb";
		echo $workbook->easy_getError();
	}
	
	// Dispose memory
	$workbook->Dispose();	
?>
