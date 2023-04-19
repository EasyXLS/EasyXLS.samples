/* ---------------------------------------------------------------
 | Tutorial 11
 |
 | This tutorial shows how to create an Excel file in C++.NET that
 | has a cell that contains SUM formula for a range of cells.
 ---------------------------------------------------------------*/

using namespace System;
using namespace EasyXLS;

int main()
{

	Console::WriteLine("Tutorial 11\n----------\n");

	// Create an instance of the class that exports Excel files
	ExcelDocument ^workbook = gcnew ExcelDocument();

	// Create a sheet
	workbook->easy_addWorksheet("Formula");

	// Get the table of data for the sheet, add data in sheet and the formula
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(workbook->easy_getSheet("Formula"));
	ExcelTable ^xlsTable = xlsFirstTab->easy_getExcelTable();
	xlsTable->easy_getCell("A1")->setValue("1");
	xlsTable->easy_getCell("A2")->setValue("2");
	xlsTable->easy_getCell("A3")->setValue("3");
	xlsTable->easy_getCell("A4")->setValue("4");
	xlsTable->easy_getCell("A6")->setValue("=SUM(A1:A4)");

	// Export Excel file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial11 - formulas in Excel.xlsx.");
	workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial11 - formulas in Excel.xlsx");

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