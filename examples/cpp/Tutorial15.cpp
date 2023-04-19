/* --------------------------------------------------------------------------------------
 * Tutorial 15
 * 
 * EasyXLS supports the following hyperlink types:
 *		1 - hyperlink to URL 
 *		2 - hyperlink to file 
 *		3 - hyperlink to UNC 
 *		4 - hyperlink to cell in the same Excel file
 *		5 - hyperlink to name 
 * 
 * The link can be placed on a range of cells.
 *
 * Every type of hyperlink accepts a tool tip description.
 *
 * Every type of hyperlink accepts a text mark. A text mark is a link inside the file.
 * ----------------------------------------------------------------------------------- */


#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 15\n----------\n");

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

			EasyXLS::IExcelWorksheetPtr xlsTab1 = (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheetAt(0);	
			EasyXLS::IExcelWorksheetPtr xlsTab2 = (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheetAt(1);	
			
			// Create hyperlink to URL
			xlsTab1->easy_addHyperlink_3(HYPERLINKTYPE_URL, "https://www.easyxls.com", "Link to URL", "B2:E2");

			// Create hyperlink to file
			xlsTab1->easy_addHyperlink_3(HYPERLINKTYPE_FILE, "c:\\myfile.xls", "Link to file", "B3");

			// Create hyperlink to UNC
			xlsTab1->easy_addHyperlink_3(HYPERLINKTYPE_UNC, "\\\\computerName\\Folder\\file.txt", "Link to UNC", "B4:D4");

			// Create hyperlink to cell on second sheet
			xlsTab1->easy_addHyperlink_3(HYPERLINKTYPE_CELL, "'Second tab'!D3", "Link to CELL", "B5");

			// Create a name on the second sheet
			xlsTab2->easy_addName_2("Name", "=Second tab!$A$1:$A$4");
			
			// Create hyperlink to name
			xlsTab1->easy_addHyperlink_3(HYPERLINKTYPE_CELL, "Name", "Link to a name", "B6");

			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial15 - hyperlinks in Excel.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial15 - hyperlinks in Excel.xlsx");
			
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