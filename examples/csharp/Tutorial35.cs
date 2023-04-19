/* --------------------------------------------------------------------
 * Tutorial 35
 * 
 * This tutorial shows how to import Excel sheet to DataSet in C#.
 * The data is imported from a specific Excel sheet (For this example 
 * we use the Excel file generated in Tutorial 09).
 * ----------------------------------------------------------------- */

using System;
using System.IO;
using System.Data;
using EasyXLS;

class Tutorial35
{
	
	[STAThread]
	static void Main()
	{
		Console.WriteLine("Tutorial 35\n-----------\n");

        // Create an instance of the class that imports Excel files
		ExcelDocument workbook = new ExcelDocument();

        // Import Excel sheet to DataSet
		Console.WriteLine("Reading file C:\\Samples\\Tutorial09.xlsx.\n");
		DataSet ds = workbook.easy_ReadXLSXSheet_AsDataSet("C:\\Samples\\Tutorial09.xlsx", "First tab");

        // Display imported DataSet values
		DataTable dt = ds.Tables[0];
		for (int row=0; row<dt.Rows.Count; row++)
			for (int column=0; column<dt.Columns.Count; column++)
				Console.WriteLine("At row " + (row + 1) + ", column " + (column + 1) +
					" the value is '" + dt.Rows[row].ItemArray[column] + "'");
 
		Console.Write("\nPress Enter to continue...");
		
		// Dispose memory
		workbook.Dispose();

		Console.ReadLine();
	}
}

