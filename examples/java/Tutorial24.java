//package testexceljava;

import EasyXLS.*;
import EasyXLS.Charts.*;

/*-------------------------------------------------------------
 | Tutorial 24
 |
 | This tutorial shows how to create an Excel file in Java and
 | to create a chart in a worksheet.
 -------------------------------------------------------------*/

public class Tutorial24{

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 24");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files
      ExcelDocument workbook = new ExcelDocument();

      // Create an worksheet
      workbook.easy_addWorksheet("SourceData");

      // Get the table of data for the worksheet
      ExcelTable xlsTable1 = ((ExcelWorksheet)workbook.easy_getSheet("SourceData")).easy_getExcelTable();

	  // Add data in cells for report header
      xlsTable1.easy_getCell(0, 0).setValue("Show Date");
      xlsTable1.easy_getCell(0, 1).setValue("Available Places");
      xlsTable1.easy_getCell(0, 2).setValue("Available Tickets");
      xlsTable1.easy_getCell(0, 3).setValue("Sold Tickets");
	  
	  // Add data in cells for chart report values
      xlsTable1.easy_getCell(1, 0).setValue("03/13/2005 00:00:00");
      xlsTable1.easy_getCell(1, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE);
      xlsTable1.easy_getCell(2, 0).setValue("03/14/2005 00:00:00");
      xlsTable1.easy_getCell(2, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE);
      xlsTable1.easy_getCell(3, 0).setValue("03/15/2005 10:00:00");
      xlsTable1.easy_getCell(3, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE);
      xlsTable1.easy_getCell(4, 0).setValue("03/16/2005 00:00:00");
      xlsTable1.easy_getCell(4, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE);

      xlsTable1.easy_getCell(1, 1).setValue("10000");
      xlsTable1.easy_getCell(2, 1).setValue("5000");
      xlsTable1.easy_getCell(3, 1).setValue("8500");
      xlsTable1.easy_getCell(4, 1).setValue("1000");

      xlsTable1.easy_getCell(1, 2).setValue("8000");
      xlsTable1.easy_getCell(2, 2).setValue("4000");
      xlsTable1.easy_getCell(3, 2).setValue("6000");
      xlsTable1.easy_getCell(4, 2).setValue("1000");

      xlsTable1.easy_getCell(1, 3).setValue("920");
      xlsTable1.easy_getCell(2, 3).setValue("1005");
      xlsTable1.easy_getCell(3, 3).setValue("342");
      xlsTable1.easy_getCell(4, 3).setValue("967");

	  // Set column widths
      xlsTable1.easy_getColumnAt(0).setWidth(100);
      xlsTable1.easy_getColumnAt(1).setWidth(100);
      xlsTable1.easy_getColumnAt(2).setWidth(100);
      xlsTable1.easy_getColumnAt(3).setWidth(100);
      
      // Create a chart
      ExcelChart xlsChart = new ExcelChart("A10", 600, 300);
      xlsChart.easy_addSeries("=SourceData!$B$1", "=SourceData!$B$2:$B$5");
      xlsChart.easy_addSeries("=SourceData!$C$1", "=SourceData!$C$2:$C$5");
      xlsChart.easy_addSeries("=SourceData!$D$1", "=SourceData!$D$2:$D$5");
      xlsChart.easy_setCategoryXAxisLabels("=SourceData!$A$2:$A$5");

      // Add the chart to the first worksheet
      ((ExcelWorksheet)workbook.easy_getSheet("SourceData")).easy_addChart(xlsChart);

      // Export Excel file
      System.out.println("Writing file C:\\Samples\\Tutorial24 - chart inside worksheet.xlsx.");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial24 - chart inside worksheet.xlsx");

      // Confirm export of Excel file
      if (workbook.easy_getError().equals(""))
        System.out.println("File successfully created.");
      else
        System.out.println("Error encountered: " + workbook.easy_getError());

      // Dispose memory
      workbook.Dispose();
    }
    catch (Exception ex) {
      ex.printStackTrace();
    }
  }
}
