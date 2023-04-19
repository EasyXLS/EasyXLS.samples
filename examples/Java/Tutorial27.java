//package testexceljava;

import EasyXLS.*;

/*-------------------------------------------------------------------------------
  | Tutorial 27
 |
 | This tutorial shows how to create an Excel file in Java and
 | encrypt the Excel file by setting the password required for opening the file.
 ------------------------------------------------------------------------------*/

public class Tutorial27 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 27");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files, having two sheets
      ExcelDocument workbook = new ExcelDocument(2);

      // Set the sheet names
      workbook.easy_getSheetAt(0).setSheetName("First tab");
      workbook.easy_getSheetAt(1).setSheetName("Second tab");

      // Set the password for protecting the Excel file when the file is open
      workbook.easy_getOptions().setPasswordToOpen("password");

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial27 - protect Excel with password and encryption.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial27 - protect Excel with password and encryption.xlsx");

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
