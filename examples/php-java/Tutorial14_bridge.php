<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*================================================================
	 | Tutorial 14
	 |
	 | This tutorial shows how to create an Excel file in PHP having
	 | a sheet and conditional formatting for cell ranges.
	  ==============================================================*/
	
	include("DataType.inc");
	include("ConditionalFormatting.inc");
	include("FontSettings.inc");
	include("Border.inc");
	include("Color.inc");

	header("Content-Type: text/html");
	
	echo "Tutorial 14<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
		
	// Create a sheet
	$workbook->easy_addWorksheet("Sheet1");

	// Get the table of data for the first worksheet
	$xlsTab = $workbook->easy_getSheet("Sheet1");	
	$xlsTable = $xlsTab->easy_getExcelTable();

	// Add data in cells
	for ($i=0; $i<6; $i++)
	{
		for ($j=0; $j<4; $j++)
		{
			if(($i<2)&&($j<2))
				$xlsTable->easy_getCell($i, $j)->setValue("12");
			else
				if(($j==2)&&($i<2))
					$xlsTable->easy_getCell($i, $j)->setValue("1000");
				else
					$xlsTable->easy_getCell($i, $j)->setValue("9");
			$xlsTable->easy_getCell($i, $j)->setDataType($DATATYPE_NUMERIC) ;
		}
	}

	// Set conditional formatting
	$xlsTab->easy_addConditionalFormatting("A1:C3", $CONDITIONALFORMATTING_OPERATOR_BETWEEN, "=9", "=11", true, true, java("java.awt.Color")->RED);

	// Set another conditional formatting
	$xlsTab->easy_addConditionalFormatting("A6:C6", $CONDITIONALFORMATTING_OPERATOR_BETWEEN, "=COS(PI())+2", "", java("java.awt.Color")->ORANGE);
	$xlsTab->easy_getConditionalFormattingAt("A6:C6")->getConditionAt(0)->setConditionType($CONDITIONALFORMATTING_CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA);
	
	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial14 - conditional formatting in Excel.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial14 - conditional formatting in Excel.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>