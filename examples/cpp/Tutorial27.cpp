/* --------------------------------------------------------------------------------
 * Tutorial 27
 *
 * This tutorial shows how to create an Excel file in C++ and
 * encrypt the Excel file by setting the password required for opening the file.
 * ----------------------------------------------------------------------------- */

#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 27\n----------\n");

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
			
			// Create two worksheets
			workbook->easy_addWorksheet_2("First tab");
			workbook->easy_addWorksheet_2("Second tab");

			// Set the password for protecting the Excel file when the file is open
			workbook->easy_getOptions()->setPasswordToOpen("password");
		
			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial27 - protect Excel with password and encryption.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial27 - protect Excel with password and encryption.xlsx");
			
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
