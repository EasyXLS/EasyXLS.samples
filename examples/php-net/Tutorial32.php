<?php
	/*=========================================================================
	 | Tutorial 32 
	 |
	 | This tutorial shows how to export data to XML Spreadsheet file in PHP.
	  =======================================================================*/
	
	include("DataType.inc");
	include("Styles.inc");

	header("Content-Type: text/html");

	echo "Tutorial 32<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");
	
	// Create two worksheets
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

	// Create an instance of the class used to format the cells
	$xlsAutoFormat = new COM("EasyXLS.ExcelAutoFormat");
	$xlsAutoFormat->InitAs($AUTOFORMAT_EASYXLS1);

	// Apply the predefined format to the cells
	$xlsFirstTable->easy_setRangeAutoFormat_2("A1:E101", $xlsAutoFormat);
	
	// Export XML Spreadsheet file
	echo "Writing file: C:\Samples\Tutorial32 - export XML spreadsheet file.xml<br>";
	$workbook->easy_WriteXMLFile_2("C:\Samples\Tutorial32 - export XML spreadsheet file.xml");
	
	// Confirm export of XML file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>
