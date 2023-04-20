<?php
	/*==================================================================
	 | Tutorial 19 
	 |
	 | This tutorial shows how to create an Excel file in PHP having
	 | multiple sheets. The first sheet is filled with data and the
	 | first cell of the second row contains data in rich text format.
	  ================================================================*/
	
	include("DataType.inc");

	header("Content-Type: text/html");
	
	echo "Tutorial 19<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");
	
	// Create two sheets
	$workbook->easy_addWorksheet_2("First tab");
	$workbook->easy_addWorksheet_2("Second tab");

	// Get the table of data for the first worksheet
	$xlsFirstTable = $workbook->easy_getSheetAt(0)->easy_getExcelTable();

    // Create the string used to set the RTF in cell
    $sFormattedValue = "This is <b>bold</b>.";
    $sFormattedValue = $sFormattedValue . "\nThis is <i>italic</i>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <u>underline</u>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <underline double>double underline</underline double>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <font color=red>red</font>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <font color=rgb(255,0,0)>red</font> too.";
    $sFormattedValue = $sFormattedValue ."\nThis is <font face=\"Arial Black\">Arial Black</font>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <font size=15pt>size 15</font>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <s>strikethrough</s>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <sup>superscript</sup>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <sub>subscript</sub>.";
    $sFormattedValue = $sFormattedValue ."\n<b>This</b> <i>is</i> <font color=red face=\"Arial Black\" size=15pt><underline double>formatted</underline double></font> <s>text</s>.";

    // Set the rich text value in cell
    $xlsFirstTable->easy_getCell(1, 0)->setHTMLValue ($sFormattedValue);
    $xlsFirstTable->easy_getCell(1, 0)->setDataType ($DATATYPE_STRING);
    $xlsFirstTable->easy_getCell(1, 0)->setWrap (True);
    $xlsFirstTable->easy_getRowAt(1)->setHeight (250);
    $xlsFirstTable->easy_getColumnAt(0)->setWidth (250);

	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial19 - RTF for Excel cells.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial19 - RTF for Excel cells.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>