/* -----------------------------------------------------------
 * Tutorial 25
 * 
 * This tutorial shows how to create an Excel file in C# and
 * to create a pivot table in a worksheet.
 * -------------------------------------------------------- */

using System;
using EasyXLS;
using EasyXLS.Constants;
using EasyXLS.PivotTables;

public class Tutorial25
{
	
	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 25\n----------\n");

        // Create an instance of the class that exports Excel files, having two sheets
		ExcelDocument workbook = new ExcelDocument(2);

        // Set the sheet names
		workbook.easy_getSheetAt(0).setSheetName("First tab");
		workbook.easy_getSheetAt(1).setSheetName("Second tab");

        //  Get the table of data for the first worksheet
		ExcelTable xlsFirstTable = ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_getExcelTable();

        // Add data in cells for report header
		xlsFirstTable.easy_getCell(0,0).setValue("Sale agent");
		xlsFirstTable.easy_getCell(0,0).setDataType(DataType.STRING);
		xlsFirstTable.easy_getCell(0,1).setValue("Sale country");
		xlsFirstTable.easy_getCell(0,1).setDataType(DataType.STRING);
		xlsFirstTable.easy_getCell(0,2).setValue("Month");
		xlsFirstTable.easy_getCell(0,2).setDataType(DataType.STRING);
		xlsFirstTable.easy_getCell(0,3).setValue("Year");
		xlsFirstTable.easy_getCell(0,3).setDataType(DataType.STRING);
		xlsFirstTable.easy_getCell(0,4).setValue("Sale amount");
		xlsFirstTable.easy_getCell(0,4).setDataType(DataType.STRING);

		xlsFirstTable.easy_getRowAt(0).setBold(true);

        // Add data in cells for report values - the source for pivot table
		xlsFirstTable.easy_getCell(1,0).setValue("John Down");
		xlsFirstTable.easy_getCell(1,1).setValue("USA");
		xlsFirstTable.easy_getCell(1,2).setValue("June");
		xlsFirstTable.easy_getCell(1,3).setValue("2010");
		xlsFirstTable.easy_getCell(1,4).setValue("550");
		
		xlsFirstTable.easy_getCell(2,0).setValue("Scott Valey");
		xlsFirstTable.easy_getCell(2,1).setValue("United Kingdom");
		xlsFirstTable.easy_getCell(2,2).setValue("June");
		xlsFirstTable.easy_getCell(2,3).setValue("2010");
		xlsFirstTable.easy_getCell(2,4).setValue("2300");
		
		xlsFirstTable.easy_getCell(3,0).setValue("John Down");
		xlsFirstTable.easy_getCell(3,1).setValue("USA");
		xlsFirstTable.easy_getCell(3,2).setValue("July");
		xlsFirstTable.easy_getCell(3,3).setValue("2010");
		xlsFirstTable.easy_getCell(3,4).setValue("3100");
		
		xlsFirstTable.easy_getCell(4,0).setValue("John Down");
		xlsFirstTable.easy_getCell(4,1).setValue("USA");
		xlsFirstTable.easy_getCell(4,2).setValue("June");
		xlsFirstTable.easy_getCell(4,3).setValue("2011");
		xlsFirstTable.easy_getCell(4,4).setValue("1050");
			
		xlsFirstTable.easy_getCell(5,0).setValue("John Down");
		xlsFirstTable.easy_getCell(5,1).setValue("USA");
		xlsFirstTable.easy_getCell(5,2).setValue("July");
		xlsFirstTable.easy_getCell(5,3).setValue("2011");
		xlsFirstTable.easy_getCell(5,4).setValue("2400");
		
		xlsFirstTable.easy_getCell(6,0).setValue("Steve Marlowe");
		xlsFirstTable.easy_getCell(6,1).setValue("France");
		xlsFirstTable.easy_getCell(6,2).setValue("June");
		xlsFirstTable.easy_getCell(6,3).setValue("2011");
		xlsFirstTable.easy_getCell(6,4).setValue("1200");
		
		xlsFirstTable.easy_getCell(7,0).setValue("Scott Valey");
		xlsFirstTable.easy_getCell(7,1).setValue("United Kingdom");
		xlsFirstTable.easy_getCell(7,2).setValue("June");
		xlsFirstTable.easy_getCell(7,3).setValue("2011");
		xlsFirstTable.easy_getCell(7,4).setValue("700");
		
		xlsFirstTable.easy_getCell(8,0).setValue("Scott Valey");
		xlsFirstTable.easy_getCell(8,1).setValue("United Kingdom");
		xlsFirstTable.easy_getCell(8,2).setValue("July");
		xlsFirstTable.easy_getCell(8,3).setValue("2011");
		xlsFirstTable.easy_getCell(8,4).setValue("360");

        // Create pivot table
		ExcelPivotTable xlsPivotTable = new ExcelPivotTable();
		xlsPivotTable.setName("Sales");
		xlsPivotTable.setSourceRange("First tab!$A$1:$E$9", workbook);
		xlsPivotTable.setLocation("A3:G15");
		xlsPivotTable.addFieldToRowLabels("Sale agent");
		xlsPivotTable.addFieldToColumnLabels("Year");
		xlsPivotTable.addFieldToValues("Sale amount","Sale amount per year",PivotTable.SUBTOTAL_SUM);  
		xlsPivotTable.addFieldToReportFilter("Sale country");
		xlsPivotTable.setOutlineForm(); 
		xlsPivotTable.setStyle(PivotTable.PIVOT_STYLE_DARK_11);

        // Add the pivot table to the second sheet
		((ExcelWorksheet)workbook.easy_getSheet("Second tab")).easy_addPivotTable(xlsPivotTable);

        // Export Excel file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial25 - pivot table in Excel.xlsx.");
        workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial25 - pivot table in Excel.xlsx");

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

