<?php
	/*=============================================================
	 | Tutorial 30 
	 |
	 | This tutorial shows how to export data to CSV file in PHP.
	  ===========================================================*/
	
	include("DataType.inc");

	header("Content-Type: text/html");

	echo "Tutorial 30<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");
	
	// Create a worksheet
	$workbook->easy_addWorksheet_2("First tab");

	// Get the table of data for the worksheet
	$xlsFirstTable = $workbook->easy_getSheetAt(0)->easy_getExcelTable();

	// Add data in cells for report header
	for ($column=0; $column<5; $column++)
	{
		$xlsFirstTable->easy_getCell(0,$column)->setValue("Column " . ($column + 1));
		$xlsFirstTable->easy_getCell(0,$column)->setDataType($DATATYPE_STRING);
	}

	// Add data in cells for report values
	for ($row=0; $row<100; $row++)
	{
		for ($column=0; $column<5; $column++)
		{
			$xlsFirstTable->easy_getCell($row+1,$column)->setValue("Data ".($row + 1).", ".($column + 1));
			$xlsFirstTable->easy_getCell($row+1,$column)->setDataType($DATATYPE_STRING);
		}
	}
	
	// Export CSV file
	echo "Writing file: C:\Samples\Tutorial30 - export CSV file.csv<br>";
	$workbook->easy_WriteCSVFile("C:\Samples\Tutorial30 - export CSV file.csv","First tab");
	
	// Confirm export of CSV file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>
