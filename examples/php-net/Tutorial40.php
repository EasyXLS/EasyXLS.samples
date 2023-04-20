<?php
	/*========================================================================
	 | Tutorial 40
	 |
	 | This tutorial shows how to convert HTML file to Excel in PHP. The
	 | HTML file generated by Tutorial 31 is imported, some data is modified
	 | and after that is exported as Excel file.
	  ======================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 40<br>";
	echo "----------<br>";

	// Create an instance of the class used to import/export Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");
	
	// Import HTML file
	echo "Reading file: C:\\Samples\\Tutorial31.html<br>";
	if ($workbook->easy_LoadHTMLFile_2("C:\\Samples\\Tutorial31.html"))
	{
		
		// Set worksheet name
		$workbook->easy_getSheetAt(0)->setSheetName("First tab");

		// Add new worksheet and add some data in cells (optional step)
		$workbook->easy_addWorksheet_2("Second tab");
		$xlsTable = $workbook->easy_getSheetAt(1)->easy_getExcelTable();
		$xlsTable->easy_getCell_2("A1")->setValue("Data added by Tutorial40");

		for ($column=0; $column<5; $column++)
		{
			$xlsTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		// Export Excel file
		echo "Writing file: C:\Samples\Tutorial40 - convert HTML to Excel.xlsx<br>";
		$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial40 - convert HTML to Excel.xlsx");
		
		// Confirm conversion of HTML to Excel
		if ($workbook->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $workbook->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial31.html";
		echo $workbook->easy_getError();
	}
	
	// Dispose memory
	$workbook->Dispose();	
?>
