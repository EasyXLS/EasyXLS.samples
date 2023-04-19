/* -----------------------------------------------------------------------------
 | Tutorial 27                                                     
 |                                                                
 | This tutorial shows how to create an Excel file in C++.NET and
 | encrypt the Excel file by setting the password required for opening the file.
 -----------------------------------------------------------------------------*/

using namespace System;
using namespace EasyXLS;

int main()
{

	Console::WriteLine("Tutorial 27\n----------\n");

	// Create an instance of the class that exports Excel files, having two sheets
	ExcelDocument ^workbook = gcnew ExcelDocument(2);
	    
	// Set the sheet names
	workbook->easy_getSheetAt(0)->setSheetName("First tab");
	workbook->easy_getSheetAt(1)->setSheetName("Second tab");

	// Set the password for protecting the Excel file when the file is open
	workbook->easy_getOptions()->setPasswordToOpen("password");

	// Export Excel file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial27 - protect Excel with password and encryption.xlsx.");
	workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial27 - protect Excel with password and encryption.xlsx");
	
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