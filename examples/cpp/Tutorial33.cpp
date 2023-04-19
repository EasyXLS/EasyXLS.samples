/* ---------------------------------------------------------------------------
 * Tutorial 33
 * 
 * This tutorial shows how to set document properties for Excel file in C++,
 * like 'Subject' property for summary information, 'Manager' property for
 * document summary information and a custom property.
 * ------------------------------------------------------------------------ */


#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 33\n----------\n");

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
			
			// Set the 'Subject' document property
			workbook->getSummaryInformation()->setSubject("This is the subject");

			// Set the 'Manager' document property
			workbook->getDocumentSummaryInformation()->setManager("This is the manager");

			// Set a custom document property
			workbook->getDocumentSummaryInformation()->setCustomProperty("PropertyName", FILEPROPERTY_VT_NUMBER, "4");
			
			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial33 - Excel file properties.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial33 - Excel file properties.xlsx");
			
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
