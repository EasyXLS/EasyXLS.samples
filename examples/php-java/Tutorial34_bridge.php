<?php require_once("http://localhost:8080/JavaBridge/java/Java.inc");

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
	$workbook = new java("EasyXLS.ExcelDocument");
	
	// Import Excel file to List
	echo "Reading file: C:\\Samples\\Tutorial09_bridge.xlsx<br><br>";
	$rows = $workbook->easy_ReadXLSXActiveSheet_AsList("C:\\Samples\\Tutorial09_bridge.xlsx");
	
echo $rows->size();
	// Confirm import of Excel file
	if ($workbook->easy_getError() == "")
	{
		// Display imported List values
		for ($rowIndex=0; $rowIndex<(int)(string)$rows->size(); $rowIndex++)
		{
			$row = $rows->elementAt($rowIndex);
			for ($cellIndex=0; $cellIndex<(int)(string)$row->size(); $cellIndex++)
			{
				echo "At row ".($rowIndex + 1).", column ".($cellIndex + 1)." the value is '".$row->elementAt($cellIndex)."'<br>";
			}
		}
	}	
	else
		echo "Error reading file C:\Samples\Tutorial09_bridge.xls " . $workbook->easy_getError();

	// Dispose memory
	$workbook->Dispose();
?>

