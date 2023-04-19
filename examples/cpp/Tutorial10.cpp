/* ----------------------------------------------------------------------------------
 * Tutorial 10
 * 
 * This tutorial shows how to export an Excel file with a merged cell range in C++.
 * ------------------------------------------------------------------------------- */


#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 10\n----------\n");

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

			// Create a worksheet
			workbook->easy_addWorksheet_2("Sheet1");

			// Get the table of data for the worksheet
			EasyXLS::IExcelWorksheetPtr xlsFirstTab= (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheet("Sheet1");
			EasyXLS::IExcelTablePtr xlsTable = xlsFirstTab->easy_getExcelTable();
			
			// Merge cells by range
			xlsTable->easy_mergeCells_2("A1:C3");

			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial10 - merge cells in Excel.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial10 - merge cells in Excel.xlsx");
			
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