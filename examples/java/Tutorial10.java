//package testexceljava;

import EasyXLS.*;

/*----------------------------------------------------------------------------------
 |Tutorial10
 | 
 | This tutorial shows how to export an Excel file with a merged cell range in Java.
  ---------------------------------------------------------------------------------*/

public class Tutorial10 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 10");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files
      ExcelDocument workbook = new ExcelDocument(1);

      // Get the table of data for the worksheet
      ExcelTable xlsTable = ((ExcelWorksheet)workbook.easy_getSheet("Sheet1")).easy_getExcelTable();

      // Merge cells by range
      xlsTable.easy_mergeCells("A1:C3");

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial10 - merge cells in Excel.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial10 - merge cells in Excel.xlsx");

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
