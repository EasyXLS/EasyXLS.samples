/* ------------------------------------------------------------------------------
| Tutorial 17													
|																	
| This tutorial shows how to create an Excel file with groups on rows in C++.NET.
| The Excel file has two worksheets. The first one is full with data and contains
| the data groups.
-------------------------------------------------------------------------------*/

#include <conio.h>
#include <stdio.h>

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{
	Console::WriteLine("Tutorial 17\n----------\n");

	// Create an instance of the class that exports Excel files having two sheets
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
	for (int row=0; row<25; row++)
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

	// Group rows and format A1:E26 cell range
	ExcelDataGroup ^xlsFirstDataGroup = gcnew ExcelDataGroup("A1:E26", DataGroup::GROUP_BY_ROWS, false);
	xlsFirstDataGroup->setAutoFormat(gcnew ExcelAutoFormat(Styles::AUTOFORMAT_EASYXLS1));
	xlsFirstTab->easy_addDataGroup(xlsFirstDataGroup );

	// Group rows and format A2:E10 cell range, outline level two, inside previous group
	ExcelDataGroup ^xlsSecondDataGroup = gcnew ExcelDataGroup("A2:E10", DataGroup::GROUP_BY_ROWS, false);		
	xlsSecondDataGroup->setAutoFormat(gcnew ExcelAutoFormat(Styles::AUTOFORMAT_EASYXLS2));		 
	xlsFirstTab->easy_addDataGroup(xlsSecondDataGroup);

	// Export Excel file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial17 - group data in Excel.xlsx.");
	workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial17 - group data in Excel.xlsx");

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