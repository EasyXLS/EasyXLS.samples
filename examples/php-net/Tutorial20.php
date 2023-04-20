<?php
	/*=========================================================
	 | Tutorial 20 
	 |
	 | This tutorial shows how to create an Excel file in PHP
	 | and apply an auto-filter to a range of cells. 
	  =======================================================*/

	include("DataType.inc");

	header("Content-Type: text/html");

	echo "Tutorial 20<br>";
	echo "----------<br>";

	// Create an instance of the class that exports Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");

	// Create a sheet
	$workbook->easy_addWorksheet_2("Sheet1");

	// Get the table of data for the worksheet
	$xlsTab = $workbook->easy_getSheet("Sheet1");
	$xlsTable = $xlsTab->easy_getExcelTable();

	// Add data in cells for report header
	for ($column=0; $column<5; $column++)
	{
		$xlsTable->easy_getCell(0,$column)->setValue("Column " . ($column + 1));
		$xlsTable->easy_getCell(0,$column)->setDataType($DATATYPE_STRING);
	}
	
	// Add data in cells for report values
	for ($row=0; $row<100; $row++)
	{
		for ($column=0; $column<5; $column++)
		{
			$xlsTable->easy_getCell($row+1,$column)->setValue("Data ".($row + 1).", ".($column + 1));
			$xlsTable->easy_getCell($row+1,$column)->setDataType($DATATYPE_STRING);
		}
	}
	
	// Apply auto-filter on cell range A1:E1
	$xlsFilter = $xlsTab->easy_getFilter();
   	$xlsFilter->setAutoFilter_2("A1:E1");
	
	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial20 - autofilter in Excel sheet.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>