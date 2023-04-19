/* -------------------------------------------------------------
 * Tutorial 31
 * 
 * This tutorial shows how to export data to HTML file in C#.
 * ---------------------------------------------------------- */

using System;
using EasyXLS;
using EasyXLS.Constants;

public class Tutorial31
{
	
	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 31\n----------\n");

        // Create an instance of the class that exports Excel files, having a sheet
		ExcelDocument workbook = new ExcelDocument(1);

        // Set the sheet name
		workbook.easy_getSheetAt(0).setSheetName("First tab");

        // Get the table of data for the worksheet
		ExcelTable xlsFirstTable = ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_getExcelTable();

        // Add data in cells for report header
		for (int column=0; column<5; column++)
		{
			xlsFirstTable.easy_getCell(0,column).setValue("Column " + (column + 1)); 
			xlsFirstTable.easy_getCell(0,column).setDataType(DataType.STRING);
		}

        // Add data in cells for report values
		for (int row=0; row<100; row++)
		{
			for (int column=0; column<5; column++)
			{
				xlsFirstTable.easy_getCell(row+1,column).setValue("Data " + (row + 1) + ", " + (column + 1)); 
				xlsFirstTable.easy_getCell(row+1,column).setDataType(DataType.STRING);
			}
		}

        // Apply a predefined format to the cells
		xlsFirstTable.easy_setRangeAutoFormat("A1:E101",new ExcelAutoFormat(Styles.AUTOFORMAT_EASYXLS1));

        // Export HTML file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial31 - export HTML file.html.");
        workbook.easy_WriteHTMLFile("C:\\Samples\\Tutorial31 - export HTML file.html", "First tab");

        // Confirm export of HTML file
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

