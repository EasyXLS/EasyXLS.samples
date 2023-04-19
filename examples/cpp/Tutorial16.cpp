/* -------------------------------------------------------------------
 * Tutorial 16
 * 
 * This tutorial shows how to create an Excel file with image in C++ 
 * The Excel file has multiple sheets.
 * The first sheet has an image inserted.
 * ---------------------------------------------------------------- */


#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 16\n----------\n");

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

			// Create two sheets 
			workbook->easy_addWorksheet_2("First tab");
			workbook->easy_addWorksheet_2("Second tab");

			// Insert image into sheet
			EasyXLS::IExcelWorksheetPtr xlsFirstTab = (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheetAt(0);
			xlsFirstTab->easy_addImage_5("C:\\Samples\\EasyXLSLogo.JPG", "A1");
			
			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial16 - images in Excel.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial16 - images in Excel.xlsx");
			
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