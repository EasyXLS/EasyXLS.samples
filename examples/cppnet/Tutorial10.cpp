/* ------------------------------------------------------------------------------------
 | Tutorial 10                                                     
 |                                                                 
 | This tutorial shows how to export an Excel file with a merged cell range in C++.NET.
  -----------------------------------------------------------------------------------*/

using namespace System;
using namespace EasyXLS;

int main()
{

	Console::WriteLine("Tutorial 10\n----------\n");

	// Create an instance of the class that exports Excel files
	ExcelDocument ^workbook = gcnew ExcelDocument(1);

	// Get the table of data for the worksheet
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(workbook->easy_getSheet("Sheet1"));
	ExcelTable ^xlsTable = xlsFirstTab->easy_getExcelTable();

	// Merge cells by range
	xlsTable->easy_mergeCells("A1:C3");    

	// Export Excel file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial10 - merge cells in Excel.xlsx.");
	workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial10 - merge cells in Excel.xlsx");

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