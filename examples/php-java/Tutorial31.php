<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*==============================================================
	 | Tutorial 31
	 |
	 | This tutorial shows how to export data to HTML file in PHP.
	  ============================================================*/
	
	include("DataType.inc");
	include("Styles.inc");

	header("Content-Type: text/html");

	echo "Tutorial 31<br>";
	echo "----------<br>";  
	
	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Create a worksheet
	$workbook->easy_addWorksheet("First tab");

	// Get the table of data for the worksheet
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
	$xlsAutoFormat = new java("EasyXLS.ExcelAutoFormat");
	$xlsAutoFormat->InitAs($AUTOFORMAT_EASYXLS1);

	// Apply the predefined format to the cells
	$xlsFirstTable->easy_setRangeAutoFormat("A1:E101", $xlsAutoFormat);
	
	// Export HTML file
	echo "Writing file: C:\Samples\Tutorial31 - export HTML file.html<br>";
	$workbook->easy_WriteHTMLFile("C:\Samples\Tutorial31 - export HTML file.html","First tab");
	
	// Confirm export of HTML file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>
