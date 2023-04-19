/* ------------------------------------------------------------------------
 * Tutorial 01
 * 
 * This code sample shows how to export DataSet to Excel file in C#.
 * The DataSet contains data from a SQL database, but it also can contain
 * data from other sources like GridView, DataGridView, DataGrid or other.
 * The cells are formatted using a predefined format.
 * --------------------------------------------------------------------- */

using System;
using System.Data;
using EasyXLS;
using EasyXLS.Constants;


public class Tutorial01
{

	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 01\n-----------\n");

        // Create an instance of the class that exports Excel files
		ExcelDocument workbook = new ExcelDocument();

        // Create the database connection
		String sConnectionString = "Initial Catalog=Northwind;Data Source=localhost;User ID=sa;Password=;";
        System.Data.SqlClient.SqlConnection sqlConnection = new System.Data.SqlClient.SqlConnection(sConnectionString);
		sqlConnection.Open();

        // Create the adapter used to fill the dataset
		String sQueryString = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', " +
										" P.ProductName AS 'Product Name', O.UnitPrice AS Price, O.Quantity , O.UnitPrice * O. Quantity AS Value" +
										" FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID";
		System.Data.SqlClient.SqlDataAdapter adp = new System.Data.SqlClient.SqlDataAdapter(sQueryString, sqlConnection);

        // Populate the dataset
		DataSet ds  = new DataSet();
		adp.Fill(ds);

        // Export the Excel file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial01 - export DataSet to Excel.xlsx.");
        workbook.easy_WriteXLSXFile_FromDataSet("c:\\Samples\\Tutorial01 - export DataSet to Excel.xlsx", ds, new ExcelAutoFormat(Styles.AUTOFORMAT_EASYXLS1), "Sheet1");

        // Confirm export of Excel file
		String sError = workbook.easy_getError();
		if (sError.Equals(""))
			Console.Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console.Write("\nError encountered: " + sError + "\nPress Enter to Exit...");

        // Close the database connection
        sqlConnection.Close();

		// Dispose memory
		workbook.Dispose();
        ds.Dispose();
        sqlConnection.Dispose();
        adp.Dispose();

		Console.ReadLine();
	}
}
