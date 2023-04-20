<?php
	/*===============================================================
	 | Tutorial 04
	 |
	 | This tutorial shows how to export data to XLSX file that has
	 | multiple sheets in PHP. The first sheet is filled with data.
	  ============================================================ */
	
	include("DataType.inc");

	header("Content-Type: text/html");

	echo "Tutorial 04<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");
	
	// Create two sheets
	$workbook->easy_addWorksheet_2("First tab");
	$workbook->easy_addWorksheet_2("Second tab");

	// Get the table of data for the first worksheet
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
	
	// Set column widths
	$xlsFirstTable->easy_getColumnAt(0)->setWidth(100);
	$xlsFirstTable->easy_getColumnAt(1)->setWidth(100);
	$xlsFirstTable->easy_getColumnAt(2)->setWidth(100);
	$xlsFirstTable->easy_getColumnAt(3)->setWidth(100);
	
	// Export the XLSX file
	echo "Writing file: C:\Samples\Tutorial04 - export data to Excel.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial04 - export data to Excel.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>
