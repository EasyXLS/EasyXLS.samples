//package testexceljava;

import EasyXLS.*;

/*---------------------------------------------------------------------
 | Tutorial 16
 |
 | This tutorial shows how to create an Excel file with image in Java.
 | The Excel file has multiple sheets.
 | The first worksheet has an image inserted.
 --------------------------------------------------------------------*/

public class Tutorial16 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 16");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files having two sheets
      ExcelDocument workbook = new ExcelDocument(2);

      // Set the sheet names
      workbook.easy_getSheetAt(0).setSheetName("First tab");
      workbook.easy_getSheetAt(1).setSheetName("Second tab");

      // Insert image into sheet
      ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_addImage("C:\\Samples\\EasyXLSLogo.JPG", "A1");

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial16 - images in Excel.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial16 - images in Excel.xlsx");

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
