/* -----------------------------------------------------------------------------
| Tutorial 36
|
| This tutorial shows how to read an Excel XLSX file in C++.NET (the
| XLSX file generated by Tutorial 04 as base template), modify
| some data and save it to another XLSX file (Tutorial36 - read XLSX file.xlsx).
----------------------------------------------------------------------------- */

using namespace System;
using namespace EasyXLS;
using namespace System::Data;
using namespace System::IO;

int main()
{
	Console::WriteLine("Tutorial 36\n----------\n");

	// Create an instance of the class that reads Excel files
	ExcelDocument ^workbook = gcnew ExcelDocument();

	// Read XLSX file
	Console::WriteLine("Reading file C:\\Samples\\Tutorial04.xlsx.\n");
	if (workbook->easy_LoadXLSXFile("C:\\Samples\\Tutorial04.xlsx")) 
	{
		// Get the table of data for the second worksheet
		ExcelWorksheet ^xlsSecondTab = safe_cast<ExcelWorksheet^>(workbook->easy_getSheet("Second tab"));
		ExcelTable ^xlsSecondTable = xlsSecondTab->easy_getExcelTable();

		// Write some data to the second sheet
        xlsSecondTable->easy_getCell("A1")->setValue("Data added by Tutorial36");
		for (int column=0; column<5; column++)
		{
			xlsSecondTable->easy_getCell(1, column)->setValue(String::Concat("Data ", (column + 1).ToString()));
		}

		// Export the new XLSX file
		Console::WriteLine("Writing file C:\\Samples\\Tutorial36 - read XLSX file.xlsx.");
		workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial36 - read XLSX file.xlsx");

		// Confirm export of Excel file
		String ^sError = workbook->easy_getError();
		if (sError->Equals(""))
			Console::Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console::Write(String::Concat("\nError encountered: ", sError, "\nPress Enter to Exit..."));
	}
    else
	{
       Console::WriteLine(String::Concat("\nError reading file C:\\Samples\\Tutorial04.xlsx \n", workbook->easy_getError(), "\nPress Enter to Exit..."));
	}
        
	// Dispose memory
    delete workbook;
	
	Console::ReadLine();

	return 0;
}