/* -------------------------------------------------------------------
 | Tutorial 03                                                     
 |                                                                
 | This tutorial shows how to create an Excel file that has 
 | multiple sheets in C++.NET. The created Excel file is empty and the 
 | next tutorial shows how to add data into sheets.                               
 -------------------------------------------------------------------*/

using namespace System;
using namespace EasyXLS;

int main()
{

	Console::WriteLine("Tutorial 03\n----------\n");

	// Create an instance of the class that creates Excel files, having two sheets
	ExcelDocument ^workbook = gcnew ExcelDocument(2);
	    
	// Set the sheet names
	workbook->easy_getSheetAt(0)->setSheetName("First tab");
	workbook->easy_getSheetAt(1)->setSheetName("Second tab");

	// Create the Excel file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial03 - create Excel file.xlsx.");
	workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial03 - create Excel file.xlsx");
	
	// Confirm the creation of Excel file
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