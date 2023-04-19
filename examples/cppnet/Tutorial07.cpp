/* -------------------------------------------------------------
 | Tutorial 07                                                     
 |                                                                
 | This code sample shows how to export an Excel file in C++.NET
 | having multiple sheet. The first sheet is filled with data
 | and the cells are formatted and locked.
 | The column header has comments.
 -------------------------------------------------------------*/

using namespace System;
using namespace System::Drawing;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{
		Console::WriteLine("Tutorial 07\n----------\n");

		// Create an instance of the class that exports Excel files, having two sheets	
		ExcelDocument ^workbook = gcnew ExcelDocument(2);
	    
		// Set the sheet names
		workbook->easy_getSheetAt(0)->setSheetName("First tab");
		workbook->easy_getSheetAt(1)->setSheetName("Second tab");

		// Protect the first sheet
		workbook->easy_getSheetAt(0)->setSheetProtected(true);

		// Get the table of data for the first worksheet
		ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(workbook->easy_getSheetAt(0));
		ExcelTable ^xlsFirstTable = xlsFirstTab->easy_getExcelTable();

		// Create the formatting style for the header
		ExcelStyle ^xlsStyleHeader = gcnew ExcelStyle("Verdana", 8, true, true, Color::Yellow);		
		xlsStyleHeader->setBackground(Color::Black);
		xlsStyleHeader->setBorderColors(Color::Gray, Color::Gray, Color::Gray, Color::Gray);
		xlsStyleHeader->setBorderStyles(Border::BORDER_MEDIUM, Border::BORDER_MEDIUM, Border::BORDER_MEDIUM, Border::BORDER_MEDIUM);	
		xlsStyleHeader->setHorizontalAlignment(Alignment::ALIGNMENT_CENTER);
		xlsStyleHeader->setVerticalAlignment(Alignment::ALIGNMENT_BOTTOM);
		xlsStyleHeader->setWrap(true);
		xlsStyleHeader->setDataType(DataType::STRING);

		// Add data in cells for report header
		for (int column=0; column<5; column++)
		{
			xlsFirstTable->easy_getCell(0,column)->setValue(String::Concat("Column ",(column + 1).ToString())); 
			xlsFirstTable->easy_getCell(0,column)->setStyle(xlsStyleHeader);

			// Add comment for report header cells
			xlsFirstTable->easy_getCell(0, column)->setComment(String::Concat("This is column no ",(column + 1).ToString()));
		}
		xlsFirstTable->easy_getRowAt(0)->setHeight(30);

		// Add data in cells for report values
		for (int row=0; row<100; row++)
		{
			for (int column=0; column<5; column++)
			{
				xlsFirstTable->easy_getCell(row+1,column)->setValue(String::Concat("Data ", (row + 1).ToString(), ", ", (column + 1).ToString())); 
			}
		}

		// Create a formatting style for cells
		ExcelStyle ^xlsStyleData = gcnew ExcelStyle();
		xlsStyleData->setHorizontalAlignment(Alignment::ALIGNMENT_LEFT);
		xlsStyleData->setForeground(Color::DarkGray);
		xlsStyleData->setWrap(false);
		xlsStyleData->setDataType(DataType::STRING);
		// Protect cells
		xlsStyleData->setLocked(true);
		xlsFirstTable->easy_setRangeStyle("A2:E101", xlsStyleData);

		// Set column widths
		xlsFirstTable->setColumnWidth(0, 70);
		xlsFirstTable->setColumnWidth(1, 100);
		xlsFirstTable->setColumnWidth(2, 70);
		xlsFirstTable->setColumnWidth(3, 100);
		xlsFirstTable->setColumnWidth(4, 70);

		// Export Excel file
		Console::WriteLine("Writing file C:\\Samples\\Tutorial07 - cell comment in Excel.xlsx.");
		workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial07 - cell comment in Excel.xlsx");

		// Confirm the export of Excel file
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