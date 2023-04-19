/* ------------------------------------------------------------------------------
 * Tutorial 17
 * 
 *  This tutorial shows how to create an Excel file with groups on rows in C++.
 * The Excel file has two worksheets. The first one is full with data and
 * contains the data groups.
 * --------------------------------------------------------------------------- */


#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 17\n----------\n");

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
			
			// Add data in cells for report header
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
			xlsFirstTable->easy_getRowAt(0)->setHeight(30);

			// Add data in cells for report values
			char*  rowNumber = (char*)malloc(sizeof(char));
			for (int row=0; row<25; row++)
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

			// Set column widths
			xlsFirstTable->setColumnWidth_2(0, 70);
			xlsFirstTable->setColumnWidth_2(1, 100);
			xlsFirstTable->setColumnWidth_2(2, 70);
			xlsFirstTable->setColumnWidth_2(3, 100);
			xlsFirstTable->setColumnWidth_2(4, 70);		

			// Group rows and format A1:E26 cell range
			EasyXLS::IExcelDataGroupPtr xlsFirstDataGroup ;
			CoCreateInstance(__uuidof(EasyXLS::ExcelDataGroup), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelDataGroupPtr), (void**) &xlsFirstDataGroup);
			xlsFirstDataGroup->setRange_2("A1:E26");
			xlsFirstDataGroup->setGroupType(DATAGROUP_GROUP_BY_ROWS);
			xlsFirstDataGroup->setCollapsed(false);
			EasyXLS::IExcelAutoFormatPtr xlsAutoFormat;
			CoCreateInstance(__uuidof(EasyXLS::ExcelAutoFormat), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelAutoFormat), (void**) &xlsAutoFormat) ;
			xlsAutoFormat->InitAs(AUTOFORMAT_EASYXLS1);
			xlsFirstDataGroup->setAutoFormat(xlsAutoFormat);
			xlsFirstTab->easy_addDataGroup(xlsFirstDataGroup );

			// Group rows and format A2:E10 cell range, outline level two, inside previous group
			EasyXLS::IExcelDataGroupPtr xlsSecondDataGroup;
			CoCreateInstance(__uuidof(EasyXLS::ExcelDataGroup), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelDataGroupPtr), (void**) &xlsSecondDataGroup);
			xlsSecondDataGroup->setRange_2("A2:E10");
			xlsSecondDataGroup->setGroupType(DATAGROUP_GROUP_BY_ROWS);
			xlsSecondDataGroup->setCollapsed(false);
			EasyXLS::IExcelAutoFormatPtr xlsAutoFormat2;
			CoCreateInstance(__uuidof(EasyXLS::ExcelAutoFormat), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelAutoFormat), (void**) &xlsAutoFormat2) ;
			xlsAutoFormat2->InitAs(AUTOFORMAT_EASYXLS2);
			xlsSecondDataGroup->setAutoFormat(xlsAutoFormat2);
			xlsFirstTab->easy_addDataGroup(xlsSecondDataGroup);

			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial17 - group data in Excel.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial17 - group data in Excel.xlsx");
			
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