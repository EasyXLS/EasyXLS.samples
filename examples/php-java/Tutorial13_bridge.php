<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*=================================================================
	 | Tutorial 13
	 |
	 | This tutorial shows how to create an Excel file in PHP having
	 | multiple sheets. The second sheet contains a named range.
	 | The A1:A10 cell range contains data validators, drop down list
	 | and whole number validation.
	  ===============================================================*/
	
	include("DataValidator.inc");

	header("Content-Type: text/html");
	
	echo "Tutorial 13<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Create two sheets
	$workbook->easy_addWorksheet("First tab");
	$workbook->easy_addWorksheet("Second tab");

	// Get the table of data for the second worksheet and populate the worksheet
	$xlsSecondTab = $workbook->easy_getSheetAt(1);
	$xlsSecondTable = $xlsSecondTab->easy_getExcelTable();
	$xlsSecondTable->easy_getCell("A1")->setValue("Range data 1");
	$xlsSecondTable->easy_getCell("A2")->setValue("Range data 2");
	$xlsSecondTable->easy_getCell("A3")->setValue("Range data 3");
	$xlsSecondTable->easy_getCell("A4")->setValue("Range data 4");

	// Create a named area range
	$xlsSecondTab->easy_addName("Range", "=Second tab!\$A\$1:\$A\$4");
	
	// Add data validation as drop down list type
	$xlsFirstTab = $workbook->easy_getSheetAt(0);
	$xlsFirstTab->easy_addDataValidator("A1:A10", $DATAVALIDATOR_VALIDATE_LIST, $DATAVALIDATOR_OPERATOR_EQUAL_TO, "=Range", "");
	
	// Add data validation as whole number type
	$xlsFirstTab->easy_addDataValidator("B1:B10", $DATAVALIDATOR_VALIDATE_WHOLE_NUMBER, $DATAVALIDATOR_OPERATOR_BETWEEN, "=4", "=100");
	
	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial13 - cell validation in Excel.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial13 - cell validation in Excel.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>