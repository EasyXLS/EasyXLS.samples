<?php
	/*======================================================================
	 | Tutorial 34 
	 |
	 | This tutorial shows how to import Excel to List in PHP. The data
	 | is imported from the active sheet of the Excel file (the Excel file 
	 | generated in Tutorial 09).
	  ====================================================================*/
	  
	header("Content-Type: text/html");
	
	echo "Tutorial 34<br>";
	echo "----------<br>";
	
	// Create an instance of the class that imports Excel files
	$workbook = new COM("EasyXLS.ExcelDocument");
	
	// Import Excel file to List
	echo "Reading file: C:\\Samples\\Tutorial09.xlsx<br><br>";
	$rows = $workbook->easy_ReadXLSXActiveSheet_AsList("C:\\Samples\\Tutorial09.xlsx");
	
	// Confirm import of Excel file
	if ($workbook->easy_getError() == "")
	{
		// Display imported List values
		for ($rowIndex=0; $rowIndex<$rows->size(); $rowIndex++)
		{
			$row = $rows->elementAt($rowIndex);
			for ($cellIndex=0; $cellIndex<$row->size(); $cellIndex++)
			{
				echo "At row ".($rowIndex + 1).", column ".($cellIndex + 1)." the value is '".$row->elementAt($cellIndex)."'<br>";
			}
		}
	}	
	else
		echo "Error reading file C:\Samples\Tutorial09.xls " . $workbook->easy_getError();

	// Dispose memory
	$workbook->Dispose();
?>

