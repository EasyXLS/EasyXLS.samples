/* ------------------------------------------------------------------
 * Tutorial 03
 *
 * This tutorial shows how to create an Excel file that has
 * multiple sheets in C++. The created Excel file is empty and the
 * next tutorial shows how to add data into sheets.
 * --------------------------------------------------------------- */

#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 03\n----------\n");

	HRESULT hr;

	// Initialize COM
	hr = CoInitialize(0);

	// Use the SUCCEEDED macro and get a pointer to the interface
	if(SUCCEEDED(hr))
	{
		// Create an instance of the class that creates Excel files
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
		
			// Create Excel file
			printf("Writing file C:\\Samples\\Tutorial03 - create Excel file.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial03 - create Excel file.xlsx");
			
			// Confirm the creation of Excel file
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