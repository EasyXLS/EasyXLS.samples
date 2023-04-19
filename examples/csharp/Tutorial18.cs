/* ----------------------------------------------------------------------
 * Tutorial 18
 * 
 * This tutorial shows how to create an Excel file in C# and 
 * freeze first row from the sheet. The Excel file has multiple sheets.
 * The first sheet is filled with data and it has a frozen row.
 * ------------------------------------------------------------------- */

using System;
using EasyXLS;
using EasyXLS.Constants;

public class Tutorial18
{
	
	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 18\n----------\n");

        // Create an instance of the class that exports Excel files having two sheets
		ExcelDocument workbook = new ExcelDocument(2);

        // Set the sheet names
		workbook.easy_getSheetAt(0).setSheetName("First tab");
		workbook.easy_getSheetAt(1).setSheetName("Second tab");

        // Get the table of data for the first worksheet
		ExcelTable xlsFirstTable = ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_getExcelTable();

        // Add data in cells for report header
		for (int column=0; column<5; column++)
		{
			xlsFirstTable.easy_getCell(0,column).setValue("Column " + (column + 1)); 
			xlsFirstTable.easy_getCell(0,column).setDataType(DataType.STRING);
		}
		xlsFirstTable.easy_getRowAt(0).setHeight(30);

        // Add data in cells for report values
		for (int row=0; row<100; row++)
		{
			for (int column=0; column<5; column++)
			{
				xlsFirstTable.easy_getCell(row+1,column).setValue("Data " + (row + 1) + ", " + (column + 1)); 
				xlsFirstTable.easy_getCell(row+1,column).setDataType(DataType.STRING);
			}
		}

        // Set column widths
		xlsFirstTable.setColumnWidth(0, 70);
		xlsFirstTable.setColumnWidth(1, 100);
		xlsFirstTable.setColumnWidth(2, 70);
		xlsFirstTable.setColumnWidth(3, 100);
		xlsFirstTable.setColumnWidth(4, 70);

        // Freeze row
		xlsFirstTable.easy_freezePanes(1, 0, 75, 0);

        // Export Excel file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial18 - freeze rows or columns in Excel.xlsx.");
        workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial18 - freeze rows or columns in Excel.xlsx");

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

