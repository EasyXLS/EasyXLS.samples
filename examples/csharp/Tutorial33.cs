/* -------------------------------------------------------------------------
 * Tutorial 33
 * 
 * This tutorial shows how to set document properties for Excel file in C#, 
 * like 'Subject' property for summary information, 'Manager' property for
 * document summary information and a custom property.
 * ---------------------------------------------------------------------- */

using System;
using EasyXLS;
using EasyXLS.Constants;

public class Tutorial33
{

	[STAThread]
	static void Main()
	{
		Console.WriteLine("Tutorial 33\n-----------\n");

        // Create an instance of the class that exports Excel files
		ExcelDocument workbook = new ExcelDocument(1);

        // Set the 'Subject' document property
		workbook.getSummaryInformation().setSubject("This is the subject");

        // Set the 'Manager' document property
		workbook.getDocumentSummaryInformation().setManager("This is the manager");

        // Set a custom document property
		workbook.getDocumentSummaryInformation().setCustomProperty("PropertyName", FileProperty.VT_NUMBER, "4");

        // Export Excel file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial33 - Excel file properties.xlsx.");
        workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial33 - Excel file properties.xlsx");

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
