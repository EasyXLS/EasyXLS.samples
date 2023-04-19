/* -----------------------------------------------------------------------
 | Tutorial 02                                                     
 |																	
 | This code sample shows how to export DataSet to Excel file in C++.NET.
 | The DataSet contains data from a SQL database, but it also can contain
 | data from other sources like GridView, DataGridView, DataGrid or other.
 | The cells are formatted using a user-defined format.
 ---------------------------------------------------------------------- */

#using <System.Xml.dll>
#using <System.dll> 

using namespace System;
using namespace System::Drawing;
using namespace System::Data;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{
	Console::WriteLine("Tutorial 02\n----------\n");

	// Create an instance of the class that exports Excel files
	ExcelDocument ^workbook = gcnew ExcelDocument();
	    
	// Create the database connection
	String ^sConnectionString = "Initial Catalog=Northwind;Data Source=localhost;User ID=sa;Password=;";
	System::Data::SqlClient::SqlConnection ^sqlConnection = gcnew System::Data::SqlClient::SqlConnection(sConnectionString);
	sqlConnection->Open();       		

	// Create the adapter used to fill the dataset
	String ^sQueryString = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', ";
	sQueryString = String::Concat	(sQueryString, " P.ProductName AS 'Product Name', O.UnitPrice AS Price, CAST(O.Quantity AS varchar) AS Quantity, O.UnitPrice * O. Quantity AS Value");
	sQueryString = String::Concat	(sQueryString, " FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID");
	System::Data::SqlClient::SqlDataAdapter ^adp = gcnew System::Data::SqlClient::SqlDataAdapter(sQueryString, sqlConnection);

	// Populate the dataset
	System::Data::DataSet ^ds  = gcnew System::Data::DataSet();
	adp->Fill(ds);

	// Create an instance of the class used to format the cells in the report
	ExcelAutoFormat ^xlsAutoFormat = gcnew ExcelAutoFormat();
    
	// Set the formatting style of the header
	ExcelStyle ^xlsHeaderStyle = gcnew ExcelStyle(Color::LightGreen);
	xlsHeaderStyle->setFontSize(12);
	xlsAutoFormat->setHeaderRowStyle(xlsHeaderStyle);

	// Set the formatting style of the cells (alternating style)
	ExcelStyle ^xlsEvenRowStripesStyle = gcnew ExcelStyle(Color::FloralWhite);
	xlsEvenRowStripesStyle->setFormat("$0.00");
	xlsEvenRowStripesStyle->setHorizontalAlignment(Alignment::ALIGNMENT_LEFT);
	xlsAutoFormat->setEvenRowStripesStyle(xlsEvenRowStripesStyle);
	ExcelStyle ^xlsOddRowStripesStyle = gcnew ExcelStyle(Color::FromArgb(240, 247, 239));
	xlsOddRowStripesStyle->setFormat("$0.00");
	xlsOddRowStripesStyle->setHorizontalAlignment (Alignment::ALIGNMENT_LEFT);
	xlsAutoFormat->setOddRowStripesStyle(xlsOddRowStripesStyle);
	ExcelStyle ^xlsLeftColumnStyle = gcnew ExcelStyle(Color::FloralWhite);
	xlsLeftColumnStyle->setFormat("mm/dd/yyyy");
	xlsLeftColumnStyle->setHorizontalAlignment(Alignment::ALIGNMENT_LEFT);
	xlsAutoFormat->setLeftColumnStyle(xlsLeftColumnStyle);

	// Export the Excel file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial02 - export DataSet to Excel with formatting.xlsx.");
	workbook->easy_WriteXLSXFile_FromDataSet("c:\\Samples\\Tutorial02 - export DataSet to Excel with formatting.xlsx", ds, xlsAutoFormat, "Sheet1");

	// Confirm export of Excel file
	String ^sError = workbook->easy_getError();
	if (sError->Equals(""))
		Console::Write("\nFile successfully created. Press Enter to Exit...");
	else
		Console::Write(String::Concat("\nError encountered: ", sError, "\nPress Enter to Exit..."));
		
	// Close the database connection.
    sqlConnection->Close();

	// Dispose memory
	delete  workbook;
    delete ds;
    delete sqlConnection;
    delete adp;

	Console::ReadLine();
	
	return 0;
}