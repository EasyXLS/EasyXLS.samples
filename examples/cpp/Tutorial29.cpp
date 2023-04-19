/* ---------------------------------------------------------------
 * Tutorial 29
 * 
 * This tutorial shows how to export data to XLSB file that has
 * multiple sheets in C++. The first sheet is filled with data.
 * ------------------------------------------------------------ */

#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 29\n----------\n");

	HRESULT hr;

	// Initialize COM
	hr = CoInitialize(0);

	// Use the SUCCEEDED macro and get a pointer to the interface
	if(SUCCEEDED(hr))
	{
		// Create a pointer to the interface that generates Excel files
		EasyXLS::IExcelDocumentPtr workbook;
		hr = CoCreateInstance(__uuidof(EasyXLS::ExcelDocument),
		NULL,
		CLSCTX_ALL,
		__uuidof(EasyXLS::IExcelDocument),
		(void**) &workbook) ;

		if(SUCCEEDED(hr)){
			
			// Create the worksheet
			workbook->easy_addWorksheet_2("First tab");
			workbook->easy_addWorksheet_2("Second tab");

			// Get the table of the first worksheet
			EasyXLS::IExcelWorksheetPtr xlsFirstTab= (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheetAt(0);
			EasyXLS::IExcelTablePtr xlsFirstTable = xlsFirstTab->easy_getExcelTable();
			
			// Add the cells for header
			char* cellValue = (char*)malloc(11*sizeof(char));
			char*  columnNumber = (char*)malloc(sizeof(char));
			for (int column=0; column<5; column++)
			{
				strcpy_s(cellValue, 8, "Column ");			
				_itoa_s(column+ 1, columnNumber, 2, 10);
				strcat_s(cellValue, 10, columnNumber);
				xlsFirstTable->easy_getCell(0,column)->setValue(cellValue); 
				xlsFirstTable->easy_getCell(0,column)->setDataType(DATATYPE_STRING);
			}

			// Add the cells for data
			char*  rowNumber = (char*)malloc(sizeof(char));
			for (int row=0; row<100; row++)
			{
				for (int column=0; column<5; column++)
				{
					strcpy_s(cellValue, 6, "Data ");	
					_itoa_s(column+ 1, columnNumber, 2, 10);
					_itoa_s(row + 1, rowNumber, 4, 10);

					strcat_s(cellValue, 10, rowNumber);
					strcat_s(cellValue, 12, ", ");
					strcat_s(cellValue, 13, columnNumber);

					xlsFirstTable->easy_getCell(row+1,column)->setValue(cellValue); 
					xlsFirstTable->easy_getCell(row+1,column)->setDataType(DATATYPE_STRING);
				}
			}
			xlsFirstTable->easy_getColumnAt(0)->setWidth(70);
			xlsFirstTable->easy_getColumnAt(1)->setWidth(100);
			xlsFirstTable->easy_getColumnAt(2)->setWidth(70);
			xlsFirstTable->easy_getColumnAt(3)->setWidth(100);
			xlsFirstTable->easy_getColumnAt(4)->setWidth(70);
		
			// Export the XLSB file
			printf("Writing file C:\\Samples\\Tutorial29 - export XLSB file.xlsb.");
			workbook->easy_WriteXLSBFile("C:\\Samples\\Tutorial29 - export XLSB file.xlsb");
			
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
