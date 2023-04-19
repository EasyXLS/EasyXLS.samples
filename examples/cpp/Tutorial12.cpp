/* ----------------------------------------------------------------
 * Tutorial 12
 * 
 * This tutorial shows how to create an Excel file in C++ having
 * multiple sheets. The second sheet contains a named area range.
 * ------------------------------------------------------------- */


#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 12\n----------\n");

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

			// Get the table of data for the second worksheet and populate the worksheet
			EasyXLS::IExcelWorksheetPtr xlsSecondTab = (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheetAt(1);
			EasyXLS::IExcelTablePtr xlsSecondTable = xlsSecondTab->easy_getExcelTable();
			xlsSecondTable->easy_getCell_2("A1")->setValue("Range data 1");
			xlsSecondTable->easy_getCell_2("A2")->setValue("Range data 2");
			xlsSecondTable->easy_getCell_2("A3")->setValue("Range data 3");
			xlsSecondTable->easy_getCell_2("A4")->setValue("Range data 4");

			// Create a named area range
			xlsSecondTab->easy_addName_2("Range", "='Second tab'!$A$1:$A$4");

			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial12 - name range in Excel.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial12 - name range in Excel.xlsx");
			
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