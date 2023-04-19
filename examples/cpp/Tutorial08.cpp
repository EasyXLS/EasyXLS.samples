/* ---------------------------------------------------------------
 * Tutorial 08
 * 
 * This tutorial shows how to create an Excel file in C++ having
 * multiple sheets. The first sheet is filled with data
 * and the cells are formatted and locked.
 * The column header has comments.
 * The first worksheet has header & footer.
 * ------------------------------------------------------------ */


#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 08\n----------\n");

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

			// Protect first sheet
			workbook->easy_getSheetAt(0)->setSheetProtected(true);

			// Get the table of data for the first worksheet
			EasyXLS::IExcelWorksheetPtr xlsFirstTab= (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheetAt(0);
			EasyXLS::IExcelTablePtr xlsFirstTable = xlsFirstTab->easy_getExcelTable();
			
			// Create the formatting style for the header
			EasyXLS::IExcelStylePtr xlsStyleHeader;
			hr = CoCreateInstance(__uuidof(EasyXLS::ExcelStyle),
			NULL,
			CLSCTX_ALL,
			__uuidof(EasyXLS::IExcelStyle),
			(void**) &xlsStyleHeader) ;

			xlsStyleHeader->setFont("Verdana");
			xlsStyleHeader->setFontSize(8);
			xlsStyleHeader->setItalic(true);
			xlsStyleHeader->setBold(true);
			xlsStyleHeader->setForeground(COLOR_YELLOW);
			xlsStyleHeader->setBackground(COLOR_BLACK);
			xlsStyleHeader->setBorderColors (COLOR_GRAY, COLOR_GRAY, COLOR_GRAY, COLOR_GRAY);
			xlsStyleHeader->setBorderStyles (BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM);
			xlsStyleHeader->setHorizontalAlignment(ALIGNMENT_ALIGNMENT_CENTER);
			xlsStyleHeader->setVerticalAlignment(ALIGNMENT_ALIGNMENT_BOTTOM);
			xlsStyleHeader->setWrap(true);
			xlsStyleHeader->setDataType(DATATYPE_STRING);

			// Add data in cells for report header
			char* cellValue = (char*)malloc(11*sizeof(char));
			char*  columnNumber = (char*)malloc(sizeof(char));
			char* commentValue = (char*)malloc(20*sizeof(char));
			for (int column=0; column<5; column++)
			{
				strcpy_s(cellValue, 8, "Column ");			
				_itoa_s(column+ 1, columnNumber, 2, 10);
				strcat_s(cellValue, 10, columnNumber);
				xlsFirstTable->easy_getCell(0,column)->setValue(cellValue); 
				xlsFirstTable->easy_getCell(0,column)->setStyle(xlsStyleHeader); 

				// Add comment for report header cells
				strcpy_s(commentValue, 19, "This is column no ");
				strcat_s(commentValue, 20, columnNumber);	
				xlsFirstTable->easy_getCell(0, column)->setComment_2(commentValue);
			}
			xlsFirstTable->easy_getRowAt(0)->setHeight(30);

			// Create a formatting style for cells
			EasyXLS::IExcelStylePtr xlsStyleData;
			hr = CoCreateInstance(__uuidof(EasyXLS::ExcelStyle),
			NULL,
			CLSCTX_ALL,
			__uuidof(EasyXLS::IExcelStyle),
			(void**) &xlsStyleData) ;

			xlsStyleData->setHorizontalAlignment(ALIGNMENT_ALIGNMENT_LEFT);
			xlsStyleData->setForeground(COLOR_DARKGRAY);
			xlsStyleData->setWrap(false);
			// Protect cells
			xlsStyleData->setLocked(true);
			xlsStyleData->setDataType(DATATYPE_STRING);

			// Add data in cells for report values
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
					xlsFirstTable->easy_getCell(row+1,column)->setStyle(xlsStyleData);
				}
			}

			// Set column widths
			xlsFirstTable->setColumnWidth_2(0, 70);
			xlsFirstTable->setColumnWidth_2(1, 100);
			xlsFirstTable->setColumnWidth_2(2, 70);
			xlsFirstTable->setColumnWidth_2(3, 100);
			xlsFirstTable->setColumnWidth_2(4, 70);			

			// Add header on center section
			xlsFirstTab->easy_getHeaderAt_2(HEADER_POSITION_CENTER)->InsertSingleUnderline();
			xlsFirstTab->easy_getHeaderAt_2(HEADER_POSITION_CENTER)->InsertFile();
			xlsFirstTab->easy_getHeaderAt_2(HEADER_POSITION_CENTER)->InsertValue(" - How to create header and footer");

			// Add header on right section
			xlsFirstTab->easy_getHeaderAt_2(HEADER_POSITION_RIGHT)->InsertDate();
			xlsFirstTab->easy_getHeaderAt_2(HEADER_POSITION_RIGHT)->InsertValue(" ");
			xlsFirstTab->easy_getHeaderAt_2(HEADER_POSITION_RIGHT)->InsertTime();

			// Add footer on center section
			xlsFirstTab->easy_getFooterAt_2(FOOTER_POSITION_CENTER)->InsertPage();
			xlsFirstTab->easy_getFooterAt_2(FOOTER_POSITION_CENTER)->InsertValue(" of ");
			xlsFirstTab->easy_getFooterAt_2(FOOTER_POSITION_CENTER)->InsertPages();

			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial08 - header and footer in Excel.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial08 - header and footer in Excel.xlsx");
			
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