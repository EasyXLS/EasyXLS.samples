<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

	/*==========================================================================
	 | Tutorial 08
	 |
	 | This tutorial shows how to create an Excel file in PHP having
	 | multiple sheets. The first sheet is filled with data
	 | and the cells are formatted and locked.
	 | The column header has comments.
	 | The first sheet has header & footer.
	  ==========================================================================*/
	
	include("DataType.inc");
	include("Alignment.inc");
	include("Border.inc");
	include("Color.inc");
	include("Header.inc");
	include("Footer.inc");

	header("Content-Type: text/html");
	
	echo "Tutorial 08<br>";
	echo "----------<br>";
	
	// Create an instance of the class that exports Excel files
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Create two sheets
	$workbook->easy_addWorksheet("First tab");
	$workbook->easy_addWorksheet("Second tab");

	// Protect first sheet
	$workbook->easy_getSheetAt(0)->setSheetProtected(true);
	
	// Get the table of data for the first worksheet
	$xlsFirstTable = $workbook->easy_getSheetAt(0)->easy_getExcelTable();

	// Create the formatting style for the header
	$xlsStyleHeader = new java("EasyXLS.ExcelStyle");
	$xlsStyleHeader->setFont("Verdana");
	$xlsStyleHeader->setFontSize(8);
	$xlsStyleHeader->setItalic(True);
	$xlsStyleHeader->setBold(True);
	$xlsStyleHeader->setForeground(java("java.awt.Color")->YELLOW);
	$xlsStyleHeader->setBackground(java("java.awt.Color")->BLACK);
	$xlsStyleHeader->setBorderColors (java("java.awt.Color")->GRAY, java("java.awt.Color")->GRAY, java("java.awt.Color")->GRAY, java("java.awt.Color")->GRAY);
	$xlsStyleHeader->setBorderStyles ($BORDER_BORDER_MEDIUM, $BORDER_BORDER_MEDIUM, $BORDER_BORDER_MEDIUM, $BORDER_BORDER_MEDIUM);
	$xlsStyleHeader->setHorizontalAlignment($ALIGNMENT_ALIGNMENT_CENTER);
	$xlsStyleHeader->setVerticalAlignment($ALIGNMENT_ALIGNMENT_BOTTOM);
	$xlsStyleHeader->setWrap(True);
	$xlsStyleHeader->setDataType($DATATYPE_STRING);
	
	// Add data in cells for report header
	for ($column=0; $column<5; $column++)
	{
		$xlsFirstTable->easy_getCell(0,$column)->setValue("Column " . ($column + 1));
		$xlsFirstTable->easy_getCell(0,$column)->setStyle($xlsStyleHeader);

		// Add comment for report header cells
		$xlsFirstTable->easy_getCell(0, $column)->setComment("This is column no " . ($column + 1));
	}
	$xlsFirstTable->easy_getRowAt(0)->setHeight(30);
	
	// Create a formatting style for cells
	$xlsStyleData = new java("EasyXLS.ExcelStyle");
	$xlsStyleData->setHorizontalAlignment($ALIGNMENT_ALIGNMENT_LEFT);
	$xlsStyleData->setForeground(java("java.awt.Color")->LIGHT_GRAY);
	$xlsStyleData->setWrap(false);
	$xlsStyleData->setLocked(true);
	$xlsStyleData->setDataType($DATATYPE_STRING);
	
	// Add data in cells for report values
	for ($row=0; $row<100; $row++)
	{
		for ($column=0; $column<5; $column++)
		{
			$xlsFirstTable->easy_getCell($row+1,$column)->setValue("Data " . ($row + 1) . ", " . ($column + 1));
			$xlsFirstTable->easy_getCell($row+1,$column)->setStyle($xlsStyleData);
		}
	}
	
	// Set column widths
	$xlsFirstTable->setColumnWidth(0, 70);
	$xlsFirstTable->setColumnWidth(1, 100);
	$xlsFirstTable->setColumnWidth(2, 70);
	$xlsFirstTable->setColumnWidth(3, 100);
	$xlsFirstTable->setColumnWidth(4, 70);
	
	// Add header on center section
	$xlsFirstTab = $workbook->easy_getSheetAt(0);
	$xlsFirstTab->easy_getHeaderAt($HEADER_POSITION_CENTER)->InsertSingleUnderline();
	$xlsFirstTab->easy_getHeaderAt($HEADER_POSITION_CENTER)->InsertFile();
	$xlsFirstTab->easy_getHeaderAt($HEADER_POSITION_CENTER)->InsertValue(" - How to create header and footer");

	// Add header on right section
	$xlsFirstTab->easy_getHeaderAt($HEADER_POSITION_RIGHT)->InsertDate();
	$xlsFirstTab->easy_getHeaderAt($HEADER_POSITION_RIGHT)->InsertValue(" ");
	$xlsFirstTab->easy_getHeaderAt($HEADER_POSITION_RIGHT)->InsertTime();

	// Add footer on center section
	$xlsFirstTab->easy_getFooterAt($FOOTER_POSITION_CENTER)->InsertPage();
	$xlsFirstTab->easy_getFooterAt($FOOTER_POSITION_CENTER)->InsertValue(" of ");
	$xlsFirstTab->easy_getFooterAt($FOOTER_POSITION_CENTER)->InsertPages();
	
	// Export Excel file
	echo "Writing file: C:\Samples\Tutorial08 - header and footer in Excel.xlsx<br>";
	$workbook->easy_WriteXLSXFile("C:\Samples\Tutorial08 - header and footer in Excel.xlsx");
	
	// Confirm export of Excel file
	if ($workbook->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $workbook->easy_getError();
		
	// Dispose memory
	$workbook->Dispose();
?>