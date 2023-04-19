/* ------------------------------------------------------------
 * Tutorial 26
 * 
 * This tutorial shows how to create an Excel file in C++ and
 * to create a pivot chart. The pivot chart is added to a
 * workshet and also to a separate chart sheet.
 * --------------------------------------------------------- */

#include "EasyXLS.h"
#include <conio.h>


int main()
{
	printf("Tutorial 26\n----------\n");

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

			// Create a chart sheet
			workbook->easy_addChart_2("Pivot chart");

			// Get the table of data for the first worksheet
			EasyXLS::IExcelWorksheetPtr xlsFirstTab= (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheetAt(0);
			EasyXLS::IExcelTablePtr xlsFirstTable = xlsFirstTab->easy_getExcelTable();
			
			// Add data in cells for report header
			xlsFirstTable->easy_getCell(0,0)->setValue("Sale agent");
			xlsFirstTable->easy_getCell(0,0)->setDataType(DATATYPE_STRING);
			xlsFirstTable->easy_getCell(0,1)->setValue("Sale country");
			xlsFirstTable->easy_getCell(0,1)->setDataType(DATATYPE_STRING);
			xlsFirstTable->easy_getCell(0,2)->setValue("Month");
			xlsFirstTable->easy_getCell(0,2)->setDataType(DATATYPE_STRING);
			xlsFirstTable->easy_getCell(0,3)->setValue("Year");
			xlsFirstTable->easy_getCell(0,3)->setDataType(DATATYPE_STRING);
			xlsFirstTable->easy_getCell(0,4)->setValue("Sale amount");
			xlsFirstTable->easy_getCell(0,4)->setDataType(DATATYPE_STRING);

			xlsFirstTable->easy_getRowAt(0)->setBold(true);

			// Add data in cells for report values - the source for pivot chart
			xlsFirstTable->easy_getCell(1,0)->setValue("John Down");
			xlsFirstTable->easy_getCell(1,1)->setValue("USA");
			xlsFirstTable->easy_getCell(1,2)->setValue("June");
			xlsFirstTable->easy_getCell(1,3)->setValue("2010");
			xlsFirstTable->easy_getCell(1,4)->setValue("550");
		
			xlsFirstTable->easy_getCell(2,0)->setValue("Scott Valey");
			xlsFirstTable->easy_getCell(2,1)->setValue("United Kingdom");
			xlsFirstTable->easy_getCell(2,2)->setValue("June");
			xlsFirstTable->easy_getCell(2,3)->setValue("2010");
			xlsFirstTable->easy_getCell(2,4)->setValue("2300");
		
			xlsFirstTable->easy_getCell(3,0)->setValue("John Down");
			xlsFirstTable->easy_getCell(3,1)->setValue("USA");
			xlsFirstTable->easy_getCell(3,2)->setValue("July");
			xlsFirstTable->easy_getCell(3,3)->setValue("2010");
			xlsFirstTable->easy_getCell(3,4)->setValue("3100");
		
			xlsFirstTable->easy_getCell(4,0)->setValue("John Down");
			xlsFirstTable->easy_getCell(4,1)->setValue("USA");
			xlsFirstTable->easy_getCell(4,2)->setValue("June");
			xlsFirstTable->easy_getCell(4,3)->setValue("2011");
			xlsFirstTable->easy_getCell(4,4)->setValue("1050");
			
			xlsFirstTable->easy_getCell(5,0)->setValue("John Down");
			xlsFirstTable->easy_getCell(5,1)->setValue("USA");
			xlsFirstTable->easy_getCell(5,2)->setValue("July");
			xlsFirstTable->easy_getCell(5,3)->setValue("2011");
			xlsFirstTable->easy_getCell(5,4)->setValue("2400");
		
			xlsFirstTable->easy_getCell(6,0)->setValue("Steve Marlowe");
			xlsFirstTable->easy_getCell(6,1)->setValue("France");
			xlsFirstTable->easy_getCell(6,2)->setValue("June");
			xlsFirstTable->easy_getCell(6,3)->setValue("2011");
			xlsFirstTable->easy_getCell(6,4)->setValue("1200");
		
			xlsFirstTable->easy_getCell(7,0)->setValue("Scott Valey");
			xlsFirstTable->easy_getCell(7,1)->setValue("United Kingdom");
			xlsFirstTable->easy_getCell(7,2)->setValue("June");
			xlsFirstTable->easy_getCell(7,3)->setValue("2011");
			xlsFirstTable->easy_getCell(7,4)->setValue("700");
		
			xlsFirstTable->easy_getCell(8,0)->setValue("Scott Valey");
			xlsFirstTable->easy_getCell(8,1)->setValue("United Kingdom");
			xlsFirstTable->easy_getCell(8,2)->setValue("July");
			xlsFirstTable->easy_getCell(8,3)->setValue("2011");
			xlsFirstTable->easy_getCell(8,4)->setValue("360");

			// Create pivot table
			EasyXLS::IExcelPivotTablePtr xlsPivotTable;
			hr = CoCreateInstance(__uuidof(EasyXLS::ExcelPivotTable),
			NULL,
			CLSCTX_ALL,
			__uuidof(EasyXLS::IExcelPivotTable),
			(void**) &xlsPivotTable);

			xlsPivotTable->setName("Sales");
			xlsPivotTable->setSourceRange("First tab!$A$1:$E$9",_variant_t((IDispatch*)workbook)); 
			xlsPivotTable->setLocation_2("A3:G15");
			xlsPivotTable->addFieldToRowLabels("Sale agent");
			xlsPivotTable->addFieldToColumnLabels("Year");
			xlsPivotTable->addFieldToValues("Sale amount","Sale amount per year",PIVOTTABLE_SUBTOTAL_SUM);  
			xlsPivotTable->addFieldToReportFilter("Sale country");
			xlsPivotTable->setOutlineForm(); 
			xlsPivotTable->setStyle(PIVOTTABLE_PIVOT_STYLE_MEDIUM_9);

			// Add the pivot table to the second sheet
			EasyXLS::IExcelWorksheetPtr xlsWorksheet = (EasyXLS::IExcelWorksheetPtr)workbook->easy_getSheet("Second tab");
			xlsWorksheet->easy_addPivotTable(xlsPivotTable);

			// Create pivot chart
			EasyXLS::IExcelPivotChartPtr xlsPivotChart1;
			hr = CoCreateInstance(__uuidof(EasyXLS::ExcelPivotChart),
			NULL,
			CLSCTX_ALL,
			__uuidof(EasyXLS::IExcelPivotChart),
			(void**) &xlsPivotChart1);

			xlsPivotChart1->setSize(600, 300);
			xlsPivotChart1->setLeftUpperCorner_2("A10");
			xlsPivotChart1->easy_setChartType(CHART_CHART_TYPE_PYRAMID_BAR);
			xlsPivotChart1->getChartTitle()->setText("Sales");
			xlsPivotChart1->setPivotTable(xlsPivotTable);
	
			// Add the pivot chart to the second sheet
			xlsWorksheet->easy_addPivotChart(xlsPivotChart1);

			// Create a clone of the pivot chart and add the clone to the chart sheet
			EasyXLS::IExcelPivotChartPtr xlsPivotChart2 = xlsPivotChart1->Clone();
			xlsPivotChart2->setSize(970, 630);
			EasyXLS::IExcelChartSheetPtr xlsChartSheet = (EasyXLS::IExcelChartSheetPtr)workbook->easy_getSheet("Pivot chart");
			xlsChartSheet->easy_setExcelChart((EasyXLS::IExcelChartPtr)xlsPivotChart2);

			// Export Excel file
			printf("Writing file C:\\Samples\\Tutorial26 - pivot chart in Excel.xlsx.");
			workbook->easy_WriteXLSXFile("C:\\Samples\\Tutorial26 - pivot chart in Excel.xlsx");
			
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