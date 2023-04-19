/* --------------------------------------------------------------
 * Tutorial 01
 *
 * This tutorial shows how to export list to Excel file in C++.
 * The list contains data from a SQL database.
 * The cells are formatted using a predefined format.
 * ----------------------------------------------------------- */

#include "EasyXLS.h"
#include <conio.h>
#import "C:\Program Files\Common Files\System\ado\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")

int main()
{

	printf("Tutorial 01\n----------\n");

	HRESULT hr ;

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

		if(SUCCEEDED(hr))
		{
			// Create the database connection
			_ConnectionPtr objConn;
			objConn.CreateInstance(__uuidof(Connection));
			objConn->Open("driver={sql server};server=(local);Database=Northwind;UID=sa;PWD=;", (BSTR) NULL, (BSTR) NULL, -1);
									
			WCHAR* sQueryString = L"SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', P.ProductName AS 'Product Name', O.UnitPrice AS Price, cast(O.Quantity AS varchar) AS Quantity , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID";
			_variant_t sqlQueryString = sQueryString ;

			// Query the database
			_RecordsetPtr objRS = NULL;
			objRS.CreateInstance(__uuidof(Recordset));
			objRS->Open( sqlQueryString, _variant_t((IDispatch*)objConn,true), adOpenStatic, adLockOptimistic, adCmdText);
			
			// Create the list that stores the query values
			EasyXLS::IListPtr lstRows;
			CoCreateInstance(__uuidof(EasyXLS::List), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IList), (void**) &lstRows) ;
		
			// Add the report header row to the list
			EasyXLS::IListPtr lstHeaderRow;
			CoCreateInstance(__uuidof(EasyXLS::List), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IList), (void**) &lstHeaderRow) ;		
			lstHeaderRow->addElement("Order Date");
			lstHeaderRow->addElement("Product Name");
			lstHeaderRow->addElement("Price");
			lstHeaderRow->addElement("Quantity");
			lstHeaderRow->addElement("Value");
			lstRows->addElement(_variant_t((IDispatch*)lstHeaderRow,true));

			VARIANT index;
			index.vt=VT_I4;
			FieldPtr field;	
			
			// Add the query values from the database to the list
			while (!(objRS->EndOfFile))
			{
				EasyXLS::IListPtr  RowList;
				CoCreateInstance(__uuidof(EasyXLS::List), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IList), (void**) &RowList) ;
				VARIANT value;					

				for (int nIndex = 0; nIndex < 5; nIndex++)
				{
					index.lVal = nIndex;
					objRS->Fields->get_Item(index, &field);
					field->get_Value (&value);
					RowList->addElement(&value);
				}			
				lstRows->addElement(_variant_t((IDispatch*)RowList,true));
						
				// Move to the next record
				objRS->MoveNext();
			}

			// Create an instance of the class used to format the cells
			EasyXLS::IExcelAutoFormatPtr xlsAutoFormat;
			CoCreateInstance(__uuidof(EasyXLS::ExcelAutoFormat), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelAutoFormat), (void**) &xlsAutoFormat) ;
			xlsAutoFormat->InitAs(AUTOFORMAT_EASYXLS1);
			
			// Export list to Excel file
			printf("Writing file C:\\Samples\\Tutorial01 - export List to Excel.xlsx.");
			hr = workbook->easy_WriteXLSXFile_FromList_2("C:\\Samples\\Tutorial01 - export List to Excel.xlsx", _variant_t((IDispatch*)lstRows,true),  _variant_t((IDispatch*)xlsAutoFormat,true), "Sheet1");				
			
			// Confirm export of Excel file
			_bstr_t sError = workbook->easy_getError();
			if (strcmp(sError, "") == 0)
			{
				printf("\nFile successfully created. Press Enter to Exit...");
			}
			else
			{
				printf("\nError encountered: %s", (LPCSTR)sError); 
			}
			
			// Close the Recordset object
			objRS->Close();
		
			// Close database connection
			objConn->Close();
						
			// Dispose memory
			workbook->Dispose();			
		}
   		else
		{
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