/* ---------------------------------------------------------------------
 * Tutorial 34
 * 
 * This tutorial shows how to import Excel to DataSet in C#. The data 
 * is imported from the active sheet of the Excel file (the Excel file 
 * generated in Tutorial 09).
 * ------------------------------------------------------------------ */

using System;
using System.IO;
using System.Data;
using EasyXLS;

public class Tutorial34
{

	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 34\n-----------\n");

        // Create an instance of the class that imports Excel files
		ExcelDocument workbook = new ExcelDocument();

        // Import Excel file to DataSet
		Console.WriteLine("Reading file C:\\Samples\\Tutorial09.xlsx.\n");
		DataSet ds = workbook.easy_ReadXLSXActiveSheet_AsDataSet("C:\\Samples\\Tutorial09.xlsx");

        // Display imported DataSet values
		DataTable dt = ds.Tables[0];
		for (int row=0; row<dt.Rows.Count; row++)
			for (int column=0; column<dt.Columns.Count; column++)
				Console.WriteLine("At row " + (row + 1) + ", column " + (column + 1) +
					" the value is '" + dt.Rows[row].ItemArray[column] + "'");

		Console.Write("\nPress Enter to Exit...");
		
		// Dispose memory
		workbook.Dispose();

		Console.ReadLine();
	}
}
