<?php
	/*======================================================================
	 | Tutorial 18 
	 |
	 | This tutorial shows how to create an Excel file in PHP and
	 | freeze first row from the sheet. The Excel file has multiple sheets.
	 | The first sheet is filled with data and it has a frozen row.
	  ====================================================================*/
	
	include("DataType.inc");

	header("Content-Type: text/html");
	
	echo "Tutorial 18<br>";
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
	$xlsFirstTable->easy_getRowAt(0)->setHeight(30);

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
	$xlsFirstTable->setColumnWidth_2(0, 70);
	$xlsFirstTable->setColumnWidth_2(1, 100);
	$xlsFirstTable->setColumnWidth_2(2, 70);
	$xlsFirstTable->setColumnWidth_2(3, 100);
	$xlsFirstTable->setColumnWidth_2(4, 70);
		
	// Freeze row
    $xlsFirstTable->easy_freezePanes_2(1, 0, 75, 0);

	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial18 - freeze rows or columns in Excel.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial18 - freeze rows or columns in Excel.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>