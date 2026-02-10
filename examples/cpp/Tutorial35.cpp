/* --------------------------------------------------------------------
 * Tutorial 35
 * 
 * This tutorial shows how to import Excel sheet to List in C++.
 * The data is imported from a specific Excel sheet (For this example
 * we use the Excel file generated in Tutorial 09).
 * ----------------------------------------------------------------- */

#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 35\n----------\n");

	HRESULT hr;

	// Initialize COM
	hr = CoInitialize(0);

	// Use the SUCCEEDED macro and get a pointer to the interface
	if(SUCCEEDED(hr))
	{
		// Create a pointer to the interface that imports Excel files
		EasyXLS::IExcelDocumentPtr workbook;
		hr = CoCreateInstance(__uuidof(EasyXLS::ExcelDocument),
		NULL,
		CLSCTX_ALL,
		__uuidof(EasyXLS::IExcelDocument),
		(void**) &workbook) ;

		if(SUCCEEDED(hr)){

			// Import Excel sheet to List
			printf("\nReading file: C:\\Samples\\Tutorial09.xlsx\n");
			EasyXLS::IListPtr  rows = workbook->easy_ReadXLSXSheet_AsList_3("C:\\Samples\\Tutorial09.xlsx", "First tab");
		
			// Confirm import of Excel file
			_bstr_t sError = workbook->easy_getError();
			if (strcmp(sError, "") == 0){
			
				// Display imported List values
				for ( int rowIndex=0; rowIndex<rows->size(); rowIndex++)
				{
					EasyXLS::IListPtr 	row = (EasyXLS::IListPtr) rows->elementAt(rowIndex);
					for (int cellIndex=0; cellIndex<row->size(); cellIndex++)
					{
						printf("At row %d, column %d the value is '%s'\n", (rowIndex+ 1), (cellIndex+ 1), (LPCSTR)((_bstr_t)row->elementAt(cellIndex)));
					}
				}
				printf("\nPress Enter to exit ...");
			}
			else
			{
				printf("\nError reading file C:\\Samples\\Tutorial09.xlsx %s\n", (LPCSTR)sError); 
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

