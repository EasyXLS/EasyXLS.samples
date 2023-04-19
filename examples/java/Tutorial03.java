//package testexceljava;

import EasyXLS.*;

/*-----------------------------------------------------------
 | Tutorial 03
 |
 | This tutorial shows how to create an Excel file that has
 | multiple sheets in Java. The created Excel file is empty
 | and the next tutorial shows how to add data into sheets.
 ----------------------------------------------------------*/

public class Tutorial03 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 03");
      System.out.println("----------");

      // Create an instance of the class that creates Excel files, having two sheets
      ExcelDocument workbook = new ExcelDocument(2);

      // Set the sheet names
      workbook.easy_getSheetAt(0).setSheetName("First tab");
      workbook.easy_getSheetAt(1).setSheetName("Second tab");

      // Create the Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial03 - create Excel file.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial03 - create Excel file.xlsx");

      // Confirm the creation of Excel file
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
