/* ----------------------------------------------------------------
| Tutorial 19													
|																	
| This tutorial shows how to create an Excel file in C++.NET having
| multiple sheets. The first sheet is filled with data and the
| first cell of the second row contains data in rich text format.
-----------------------------------------------------------------*/

#include <conio.h>
#include <stdio.h>

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{
	Console::WriteLine("Tutorial 19\n----------\n");

	// Create an instance of the class that exports Excel files having two sheets
	ExcelDocument ^workbook = gcnew ExcelDocument(2);
	
	// Set the sheet names
	workbook->easy_getSheetAt(0)->setSheetName("First tab");
	workbook->easy_getSheetAt(1)->setSheetName("Second tab");
	
	// Get the table of data for the first worksheet
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(workbook->easy_getSheetAt(0));
	ExcelTable ^xlsFirstTable = xlsFirstTab->easy_getExcelTable();

	// Create the string used to set the RTF in cell
	String ^sFormattedValue = "This is <b>bold</b>.";
	sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <i>italic</i>.");
	sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <u>underline</u>.");
	sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <underline double>double underline</underline double>.");
	sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <font color=red>red</font>.");
	sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <font color=rgb(255,0,0)>red</font> too.");
	sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <font face=\"Arial Black\">Arial Black</font>.");
	sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <font size=15pt>size 15</font>.");
	sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <s>strikethrough</s>.");
	sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <sup>superscript</sup>.");
	sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <sub>subscript</sub>.");
	sFormattedValue = String::Concat(sFormattedValue ,"\n<b>This</b> <i>is</i> <font color=red face=\"Arial Black\" size=15pt><underline double>formatted</underline double></font> <s>text</s>.");

	// Set the rich text value in cell
	xlsFirstTable->easy_getCell(1, 0)->setHTMLValue(sFormattedValue); 
	xlsFirstTable->easy_getCell(1, 0)->setDataType(DataType::STRING);
	xlsFirstTable->easy_getCell(1, 0)->setWrap(true); 
	xlsFirstTable->easy_getRowAt(1)->setHeight(250);
	xlsFirstTable->easy_getColumnAt(0)->setWidth(250);

	// Export Excel file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial19 - RTF for Excel cells.xlsx.");
	workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial19 - RTF for Excel cells.xlsx");

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