/* ---------------------------------------------------------------------
| Tutorial 16                                                    
|                                                                
| This tutorial shows how to create an Excel file with image in C++.NET.
| The Excel file has multiple sheets.
| The first worksheet has an image inserted.
--------------------------------------------------------------------- */

#include <conio.h>
#include <stdio.h>

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{
	Console::WriteLine("Tutorial 16\n----------\n");

	// Create an instance of the class that exports Excel files having two sheets
	ExcelDocument ^workbook = gcnew ExcelDocument(2);
	
	// Set the sheet names
	workbook->easy_getSheetAt(0)->setSheetName("First tab");
	workbook->easy_getSheetAt(1)->setSheetName("Second tab");
	
	// Insert image into sheet
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(workbook->easy_getSheetAt(0));
	xlsFirstTab->easy_addImage("C:\\Samples\\EasyXLSLogo.JPG", "A1");

	// Export Excel file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial16 - images in Excel.xlsx.");
	workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial16 - images in Excel.xlsx");
	
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