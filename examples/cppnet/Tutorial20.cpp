/* ---------------------------------------------------------
| Tutorial 20													
|																	
| This tutorial shows how to create an Excel file in C++.NET
| and apply an auto-filter to a range of cells.
----------------------------------------------------------*/

#include <conio.h>
#include <stdio.h>

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{
	Console::WriteLine("Tutorial 20\n----------\n");

	// Create an instance of the class that exports Excel files having one sheet
	ExcelDocument ^workbook = gcnew ExcelDocument(1);
	
	// Get the table of data for the worksheet
	ExcelWorksheet ^xlsTab = safe_cast<ExcelWorksheet^>(workbook->easy_getSheet("Sheet1"));
	ExcelTable ^xlsTable = xlsTab->easy_getExcelTable();

	// Add data in cells for report header
	for (int column=0; column<5; column++)
	{
		xlsTable->easy_getCell(0,column)->setValue(String::Concat("Column ",(column + 1).ToString())); 
		xlsTable->easy_getCell(0,column)->setDataType(DataType::STRING);
	}

	// Add data in cells for report values
	for (int row=0; row<100; row++)
	{
		for (int column=0; column<5; column++)
		{
			xlsTable->easy_getCell(row+1,column)->setValue(String::Concat("Data ", (row + 1).ToString(), ", ", (column + 1).ToString())); 
			xlsTable->easy_getCell(row+1,column)->setDataType(DataType::STRING);
		}
	}

	// Apply auto-filter on cell range A1:E1
	ExcelFilter ^xlsFilter = xlsTab->easy_getFilter();
	xlsFilter->setAutoFilter("A1:E1");

	// Export Excel file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial20 - autofilter in Excel sheet.xlsx.");
	workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial20 - autofilter in Excel sheet.xlsx");

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