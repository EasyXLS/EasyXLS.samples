<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*==============================================================
	 | Tutorial 28
	 |
	 | This tutorial shows how to export data to XLS file that has
	 | multiple sheets in PHP. The first sheet is filled with data.
	  =============================================================*/
	
	include("DataType.inc");

	header("Content-Type: text/html");

	echo "Tutorial 28<br>";
	echo "----------<br>";
	

	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Create two sheets
	$workbook->easy_addWorksheet("First tab");
	$workbook->easy_addWorksheet("Second tab");

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
	$xlsFirstTable->setColumnWidth(0, 70);
	$xlsFirstTable->setColumnWidth(1, 100);
	$xlsFirstTable->setColumnWidth(2, 70);
	$xlsFirstTable->setColumnWidth(3, 100);
	$xlsFirstTable->setColumnWidth(4, 70);
	
	// Export the XLS file
	echo "Writing file: C:\Samples\Tutorial28 - export XLS file.xls<br>";
	$workbook->easy_WriteXLSFile("C:\Samples\Tutorial28 - export XLS file.xls");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>