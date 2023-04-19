/* -------------------------------------------------------------
 * Tutorial 11
 * 
 * This tutorial shows how to create an Excel file in C++ 
 * has a cell that contains SUM formula for a range of cells.
 * ---------------------------------------------------------- */


#include "EasyXLS.h"
#include <conio.h>

int main()
{
	printf("Tutorial 11\n----------\n");

	HRESULT hr;

	// Initialize COM
	hr = CoInitialize(0);

	// Use the SUCCEEDED macro and get a pointer to the interface
	if(SUCCEEDED(hr))
	{
		// Create a pointer to the interface that exports Excel files
		EasyXLS::IExcelDocumentPtr workbook;
		hr = CoCreateInstance(__uuidof(EasyXLS::ExcelDocument),
		NULL,
		CLSCTX_ALL,
		__uuidof(EasyXLS::IExcelDocument),
		(void**) &workbook) ;

		if(SUCCEEDED(hr)){

			// Create a sheet
			workbook->easy_addWorksheet_2("Formula");

			// Get the table of data for the sheet, add data in sheet and the formula
			EasyXLS::IExcelWorksheetPtr xlsFirstTab= (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheet("Formula");
			EasyXLS::IExcelTablePtr xlsTable = xlsFirstTab->easy_getExcelTable();
			xlsTable->easy_getCell_2("A1")->setValue("1");
			xlsTable->easy_getCell_2("A2")->setValue("2");
			xlsTable->easy_getCell_2("A3")->setValue("3");
			xlsTable->easy_getCell_2("A4")->setValue("4");
			xlsTable->easy_getCell_2("A6")->setValue("=SUM(A1:A4)");

			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial11 - formulas in Excel.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial11 - formulas in Excel.xlsx");
			
			// Confirm export of Excel file
			_bstr_t sError = workbook->easy_getError();
			if (strcmp(sError, "") == 0){
				printf("\nFile successfully created. Press Enter to Exit...");
			}
			else{
				printf("\nError encountered: %s", (LPCSTR)sError); 
			}
			
			// Dispose memory
			workbook->Dispose();
		}
		else{
			printf("Object is not available!");
		}
	}
	else{
		printf("COM can't be initialized!");
	}

	// Uninitialize COM
	CoUninitialize();

	_getch();
	return 0;
}