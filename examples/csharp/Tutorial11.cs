/* ----------------------------------------------------------------
 * Tutorial 11
 * 
 * This tutorial shows how to create an Excel file in C# that
 * has a cell that contains SUM formula for a range of cells.
 * ----------------------------------------------------------------- */

using System;
using EasyXLS;

public class Tutorial11
{

	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 11\n----------\n");

        // Create an instance of the class that exports Excel files
        ExcelDocument workbook = new ExcelDocument();

        // Create a sheet
		workbook.easy_addWorksheet("Formula");

        // Get the table of data for the sheet, add data in sheet and the formula
		ExcelTable xlsTable = ((ExcelWorksheet)workbook.easy_getSheet("Formula")).easy_getExcelTable();
		xlsTable.easy_getCell("A1").setValue("1");
		xlsTable.easy_getCell("A2").setValue("2");
		xlsTable.easy_getCell("A3").setValue("3");
		xlsTable.easy_getCell("A4").setValue("4");
		xlsTable.easy_getCell("A6").setValue("=SUM(A1:A4)");

        // Export Excel file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial11 - formulas in Excel.xlsx.");
        workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial11 - formulas in Excel.xlsx");

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


