/* ---------------------------------------------------------------------------------
 * Tutorial 10
 * 
 * This tutorial shows how to export an Excel file with a merged cell range in C#.
 * ------------------------------------------------------------------------------ */

using System;
using EasyXLS;

public class Tutorial10
{

	[STAThread]
	static void Main()
	{
		Console.WriteLine("Tutorial 10\n-----------\n");

        // Create an instance of the class that exports Excel files
		ExcelDocument workbook = new ExcelDocument(1);

        // Get the table of data for the worksheet
		ExcelTable xlsTable = ((ExcelWorksheet)workbook.easy_getSheet("Sheet1")).easy_getExcelTable();

        // Merge cells by range
		xlsTable.easy_mergeCells("A1:C3");

        // Export Excel file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial10 - merge cells in Excel.xlsx.");
        workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial10 - merge cells in Excel.xlsx");

        // Confirm export of Excel file
		String sError = workbook.easy_getError();
		if (sError.Equals(""))
			Console.Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console.Write("\nError encountered: " + sError + "\nPress Enter to Exit...");
		
		// Dispose memory
		workbook.Dispose();

		Console.ReadLine();
	}
}
