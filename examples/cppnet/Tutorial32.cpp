/* --------------------------------------------------------------------------
 | Tutorial 32                                                     
 |                                                                
 | This tutorial shows how to export data to XML Spreadsheet file in C++.NET.
 ------------------------------------------------------------------------- */

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{
	Console::WriteLine("Tutorial 32\n----------\n");

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

	// Add data in cells for report values
	for (int row=0; row<100; row++)
	{
		for (int column=0; column<5; column++)
		{
			xlsFirstTable->easy_getCell(row+1,column)->setValue(String::Concat("Data ", (row + 1).ToString(), ", ", (column + 1).ToString())); 
			xlsFirstTable->easy_getCell(row+1,column)->setDataType(DataType::STRING);
		}
	}

	// Apply a predefined format to the cells
	xlsFirstTable->easy_setRangeAutoFormat("A1:E101", gcnew ExcelAutoFormat(Styles::AUTOFORMAT_EASYXLS1));

	// Export XML Spreadsheet file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial32 - export XML spreadsheet file.xml.");
	workbook->easy_WriteXMLFile("C:\\Samples\\Tutorial32 - export XML spreadsheet file.xml");

	// Confirm export of XML file
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