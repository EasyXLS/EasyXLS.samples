<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*=============================================================================
	 | Tutorial 17
	 |
	 | This tutorial shows how to create an Excel file with groups on rows in PHP.
	 | The Excel file has two worksheets. The first one is full with data and
	 | contains the data groups.
	  ===========================================================================*/
	
	include("DataType.inc");
	include("Styles.inc");
	include("DataGroup.inc");

	header("Content-Type: text/html");
	
	echo "Tutorial 17<br>";
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
	for ($row=0; $row<25; $row++)
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
		
	// Group rows and format A1:E26 cell range
    $xlsFirstDataGroup = new java("EasyXLS.ExcelDataGroup");
    $xlsFirstDataGroup->setRange("A1:E26");
    $xlsFirstDataGroup->setGroupType ($DATAGROUP_GROUP_BY_ROWS);
    $xlsFirstDataGroup->setCollapsed (False);
    $xlsAutoFormat = new java("EasyXLS.ExcelAutoFormat");
    $xlsAutoFormat->InitAs($AUTOFORMAT_EASYXLS1);
    $xlsFirstDataGroup->setAutoFormat($xlsAutoFormat);
    $workbook->easy_getSheetAt(0)->easy_addDataGroup($xlsFirstDataGroup);

    // Group rows and format A2:E10 cell range, outline level two, inside previous group
    $xlsSecondDataGroup = new java("EasyXLS.ExcelDataGroup");
    $xlsSecondDataGroup->setRange("A2:E10");
    $xlsSecondDataGroup->setGroupType($DATAGROUP_GROUP_BY_ROWS);
    $xlsSecondDataGroup->setCollapsed(False);
    $xlsAutoFormat2 = new java("EasyXLS.ExcelAutoFormat");
    $xlsAutoFormat2->InitAs($AUTOFORMAT_EASYXLS2);
    $xlsSecondDataGroup->setAutoFormat($xlsAutoFormat2);
    $workbook->easy_getSheetAt(0)->easy_addDataGroup($xlsSecondDataGroup);
	
	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial17 - group data in Excel.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial17 - group data in Excel.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>