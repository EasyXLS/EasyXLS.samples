//package testexceljava;

import EasyXLS.*;

/*-------------------------------------------------------------
 | Tutorial 11
 | 
 | This tutorial shows how to create an Excel file in Java that
 | has a cell that contains SUM formula for a range of cells.
 -------------------------------------------------------------*/

public class Tutorial11 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 11");
      System.out.println("----------");

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
      System.out.println("Writing file: C:\\Samples\\Tutorial11 - formulas in Excel.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial11 - formulas in Excel.xlsx");

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
