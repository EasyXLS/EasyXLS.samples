/* ----------------------------------------------------------------
 | Tutorial 28                                                     
 |                                                                
 | This tutorial shows how to export data to XLS file that has
 | multiple sheets in C++.NET. The first sheet is filled with data.   
 ----------------------------------------------------------------*/

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{

	Console::WriteLine("Tutorial 28\n----------\n");

	// Create an instance of the class that exports Excel files, having two sheets	
	ExcelDocument ^workbook = gcnew ExcelDocument(2);
	    
	// Set the sheet names
	workbook->easy_getSheetAt(0)->setSheetName("First tab");
	workbook->easy_getSheetAt(1)->setSheetName("Second tab");

	// Get the table of data for the first worksheet
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(workbook->easy_getSheetAt(0));
	ExcelTable ^xlsFirstTable = xlsFirstTab->easy_getExcelTable();

	// Add data in cells for report header
	for (int column=0; column<5; column++)
	{
		xlsFirstTable->easy_getCell(0,column)->setValue(String::Concat("Column ",(column + 1).ToString())); 
		xlsFirstTable->easy_getCell(0,column)->setDataType(DataType::STRING);
	}
	xlsFirstTable->easy_getRowAt(0)->setHeight(30);

	// Add data in cells for report values
	for (int row=0; row<100; row++)
	{
		for (int column=0; column<5; column++)
		{
			xlsFirstTable->easy_getCell(row+1,column)->setValue(String::Concat("Data ", (row + 1).ToString(), ", ", (column + 1).ToString())); 
			xlsFirstTable->easy_getCell(row+1,column)->setDataType(DataType::STRING);
		}
	}

	// Set column widths
	xlsFirstTable->setColumnWidth(0, 70);
	xlsFirstTable->setColumnWidth(1, 100);
	xlsFirstTable->setColumnWidth(2, 70);
	xlsFirstTable->setColumnWidth(3, 100);
	xlsFirstTable->setColumnWidth(4, 70);

	// Export the XLS file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial28 - export XLS file.xls.");
	workbook->easy_WriteXLSFile("C:\\Samples\\Tutorial28 - export XLS file.xls");
	
	// Confirm export of Excel file
	String ^sError = workbook->easy_getError();
	if (sError->Equals(""))
		Console::Write("\nFile successfully created. Press Enter to Exit...");
	else
		Console::Write(String::Concat("\nError encountered: ", sError, "\nPress Enter to Exit..."));

	// Dispose memory
	delete workbook;

	Console::ReadLine();

	return 0;
}