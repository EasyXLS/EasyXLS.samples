/* ------------------------------------------------------------------
 * Tutorial 19
 * 
 * This tutorial shows how to create an Excel file in C++ having
 * multiple sheets. The first sheet is filled with data and the
 * first cell of the second row contains data in rich text format.
 * --------------------------------------------------------------- */


#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 19\n----------\n");

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

			// Get the table of data for the first worksheet
			EasyXLS::IExcelWorksheetPtr xlsFirstTab= (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheetAt(0);
			EasyXLS::IExcelTablePtr xlsFirstTable = xlsFirstTab->easy_getExcelTable();
			
			// Create the string used to set the RTF in cell
			char* sFormattedValue = (char*)malloc(536*sizeof(char));
			strcpy_s(sFormattedValue, 21, "This is <b>bold</b>.");		
			strcpy_s(sFormattedValue , 24, "\nThis is <i>italic</i>.");
			strcpy_s(sFormattedValue , 27, "\nThis is <u>underline</u>.");
			strcpy_s(sFormattedValue , 64, "\nThis is <underline double>double underline</underline double>.");
			strcpy_s(sFormattedValue , 37, "\nThis is <font color=red>red</font>.");
			strcpy_s(sFormattedValue , 50, "\nThis is <font color=rgb(255,0,0)>red</font> too.");
			strcpy_s(sFormattedValue , 56, "\nThis is <font face=\"Arial Black\">Arial Black</font>.");
			strcpy_s(sFormattedValue , 41, "\nThis is <font size=15pt>size 15</font>.");
			strcpy_s(sFormattedValue , 31, "\nThis is <s>strikethrough</s>.");
			strcpy_s(sFormattedValue , 33, "\nThis is <sup>superscript</sup>.");
			strcpy_s(sFormattedValue , 31, "\nThis is <sub>subscript</sub>.");
			strcpy_s(sFormattedValue , 138, "\n<b>This</b> <i>is</i> <font color=red face=\"Arial Black\" size=15pt> <underline double>formatted</underline double></font> <s>text</s>.");

			// Set the rich text value in cell
			xlsFirstTable->easy_getCell(1, 0)->setHTMLValue(sFormattedValue); 
			xlsFirstTable->easy_getCell(1, 0)->setDataType(DATATYPE_STRING);
			xlsFirstTable->easy_getCell(1, 0)->setWrap(true); 
			xlsFirstTable->easy_getRowAt(1)->setHeight(250);
			xlsFirstTable->easy_getColumnAt(0)->setWidth(250);

			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial19 - RTF for Excel cells.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial19 - RTF for Excel cells.xlsx");
			
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