/* ----------------------------------------------------------
 | Tutorial 21                                                     
 |                                                                 
 | This tutorial shows how to create an Excel file in C++.NET
 | having a worksheet and a chart sheet.
  ---------------------------------------------------------*/

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Charts;
using namespace EasyXLS::Constants;

int main()
{
	Console::WriteLine("Tutorial 21\n----------\n");

	// Create an instance of the class that exports Excel files
	ExcelDocument ^workbook = gcnew ExcelDocument();

	// Create a worksheet
	workbook->easy_addWorksheet("SourceData");

	// Get the table of data for the worksheet
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(workbook->easy_getSheet("SourceData"));
	ExcelTable ^xlsTable1 = xlsFirstTab->easy_getExcelTable();

	// Add data in cells for report header
	xlsTable1->easy_getCell(0, 0)->setValue("Show Date");
	xlsTable1->easy_getCell(0, 1)->setValue("Available Places");
	xlsTable1->easy_getCell(0, 2)->setValue("Available Tickets");
	xlsTable1->easy_getCell(0, 3)->setValue("Sold Tickets");

	// Add data in cells for chart report values
	xlsTable1->easy_getCell(1, 0)->setValue("03/13/2005 00:00:00");
	xlsTable1->easy_getCell(1, 0)->setFormat(EasyXLS::Constants::Format::FORMAT_DATE);
	xlsTable1->easy_getCell(2, 0)->setValue("03/14/2005 00:00:00");
	xlsTable1->easy_getCell(2, 0)->setFormat(EasyXLS::Constants::Format::FORMAT_DATE);
	xlsTable1->easy_getCell(3, 0)->setValue("03/15/2005 00:00:00");
	xlsTable1->easy_getCell(3, 0)->setFormat(EasyXLS::Constants::Format::FORMAT_DATE);
	xlsTable1->easy_getCell(4, 0)->setValue("03/16/2005 00:00:00");
	xlsTable1->easy_getCell(4, 0)->setFormat(EasyXLS::Constants::Format::FORMAT_DATE);

	xlsTable1->easy_getCell(1, 1)->setValue("10000");
	xlsTable1->easy_getCell(2, 1)->setValue("5000");
	xlsTable1->easy_getCell(3, 1)->setValue("8500");
	xlsTable1->easy_getCell(4, 1)->setValue("1000");

	xlsTable1->easy_getCell(1, 2)->setValue("8000");
	xlsTable1->easy_getCell(2, 2)->setValue("4000");
	xlsTable1->easy_getCell(3, 2)->setValue("6000");
	xlsTable1->easy_getCell(4, 2)->setValue("1000");

	xlsTable1->easy_getCell(1, 3)->setValue("920");
	xlsTable1->easy_getCell(2, 3)->setValue("1005");
	xlsTable1->easy_getCell(3, 3)->setValue("342");
	xlsTable1->easy_getCell(4, 3)->setValue("967");

	// Set column widths
	xlsTable1->easy_getColumnAt(0)->setWidth(100);
	xlsTable1->easy_getColumnAt(1)->setWidth(100);
	xlsTable1->easy_getColumnAt(2)->setWidth(100);
	xlsTable1->easy_getColumnAt(3)->setWidth(100);

	// Add a chart sheet
	workbook->easy_addChart("Chart", "=SourceData!$A$1:$D$5", Chart::SERIES_IN_COLUMNS);

	// Export Excel file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial21 - chart sheet in Excel.xlsx.");
	workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial21 - chart sheet in Excel.xlsx");

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