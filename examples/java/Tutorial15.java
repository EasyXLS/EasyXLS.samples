//package testexceljava;

import EasyXLS.*;

/*--------------------------------------------------------------------------------------
 | Tutorial 15
 |
 | This tutorial shows how to create an Excel file with hyperlinks in Java.
 |
 | EasyXLS supports the following hyperlink types:
 | (1) - hyperlink to URL
 | (2) - hyperlink to file
 | (3) - hyperlink to UNC
 | (4) - hyperlink to cell in the same Excel file
 | (5) - hyperlink to name
 |
 | Every type of hyperlink accepts a tool tip description.
 |
 | Every type of hyperlink accepts a text mark. A text mark is a link inside the file.
  ------------------------------------------------------------------------------------*/

public class Tutorial15 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 15");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files having two sheets
      ExcelDocument workbook = new ExcelDocument(2);

      // Set the sheet names
      ExcelWorksheet xlsTab1 = (ExcelWorksheet) workbook.easy_getSheetAt(0);
      ExcelWorksheet xlsTab2 = (ExcelWorksheet) workbook.easy_getSheetAt(1);
      xlsTab1.setSheetName("First tab");
      xlsTab2.setSheetName("Second tab");

      // Create hyperlink to URL
      xlsTab1.easy_addHyperlink(EasyXLS.Constants.HyperlinkType.URL, "http://www.euoutsourcing.com", "Link to URL", "B2:E2");

      // Create hyperlink to file
      xlsTab1.easy_addHyperlink(EasyXLS.Constants.HyperlinkType.FILE, "c:\\myfile.xls", "Link to file", "B3");

      // Create hyperlink to UNC
      xlsTab1.easy_addHyperlink(EasyXLS.Constants.HyperlinkType.UNC, "\\\\computerName\\Folder\\file.txt", "Link to UNC", "B4:D4");

      // Create hyperlink to cell on second sheet
      xlsTab1.easy_addHyperlink(EasyXLS.Constants.HyperlinkType.CELL, "'Second tab'!D3", "Link to CELL", "B5");

      // Create a name on the second sheet
      xlsTab2.easy_addName("Name", "=Second tab!$A$1:$A$4");

      // Create hyperlink to name
      xlsTab1.easy_addHyperlink(EasyXLS.Constants.HyperlinkType.CELL, "Name", "Link to a name", "B6");

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial15 - hyperlinks in Excel.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial15 - hyperlinks in Excel.xlsx");

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
