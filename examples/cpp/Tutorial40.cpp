/* -----------------------------------------------------------------------
 * Tutorial 40
 * 
 * This tutorial shows how to convert HTML file to Excel in C++. The
 * HTML file generated by Tutorial 31 is imported, some data is modified
 * and after that is exported as Excel file.
 * -------------------------------------------------------------------- */

#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 40\n----------\n");

	HRESULT hr;

	// Initialize COM
	hr = CoInitialize(0);

	// Use the SUCCEEDED macro and get a pointer to the interface
	if(SUCCEEDED(hr))
	{
		// Create a pointer to the interface used to import/export Excel files
		EasyXLS::IExcelDocumentPtr workbook;
		hr = CoCreateInstance(__uuidof(EasyXLS::ExcelDocument),
		NULL,
		CLSCTX_ALL,
		__uuidof(EasyXLS::IExcelDocument),
		(void**) &workbook) ;

		if(SUCCEEDED(hr)){

			// Import HTML file	
			printf("\nReading file: C:\\Samples\\Tutorial31.html\n");
			if (workbook->easy_LoadHTMLFile_2("C:\\Samples\\Tutorial31.html"))
			{
				// Set worksheet name
				workbook->easy_getSheetAt(0)->setSheetName("First tab");

				// Add new worksheet and add some data in cells (optional step)
				workbook->easy_addWorksheet_2("Second tab"); 
				EasyXLS::IExcelWorksheetPtr xlsSecondTab = (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheetAt(1);
				EasyXLS::IExcelTablePtr xlsTable = xlsSecondTab->easy_getExcelTable();
				xlsTable->easy_getCell_2("A1")->setValue("Data added by Tutorial40");

				char* cellValue = (char*)malloc(11*sizeof(char));
				char*  columnNumber = (char*)malloc(sizeof(char));
				for (int column=0; column<5; column++)
				{
					strcpy_s(cellValue, 6, "Data ");			
					_itoa_s(column+ 1, columnNumber, 2, 10);
					strcat_s(cellValue, 10, columnNumber);
					xlsTable->easy_getCell(1, column)->setValue(cellValue);
				}
			
				
				// Export Excel file
				printf("Writing file C:\\Samples\\Tutorial40 - convert HTML to Excel.xlsx.");
				workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial40 - convert HTML to Excel.xlsx");
				
				// Confirm conversion of HTML to Excel
				_bstr_t sError = workbook->easy_getError();
				if (strcmp(sError, "") == 0){
					printf("\nFile successfully created. Press Enter to Exit...");
				}
				else{
					printf("\nError encountered: %s", (LPCSTR)sError); 
				}
			}
			else
			{
				printf("\nError reading file C:\\Samples\\Tutorial31.html %s\n", (LPCSTR)((_bstr_t)workbook->easy_getError())); 
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
