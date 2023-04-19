//package testexceljava;

/*-----------------------------------------------------------------------
 | Tutorial 40
 |
 | This tutorial shows how to convert HTML file to Excel in Java. The
 | HTML file generated by Tutorial 31 is imported, some data is modified
 | and after that is exported as Excel file.
 ----------------------------------------------------------------------*/

import java.io.FileInputStream;
import EasyXLS.*;

public class Tutorial40 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 40");
      System.out.println("----------");

      // Create an instance of the class used to import/export Excel files
      ExcelDocument workbook = new ExcelDocument();

      // Import HTML file
      System.out.println("Reading file C:\\Samples\\Tutorial31.html.");
      FileInputStream file = new FileInputStream("C:\\Samples\\Tutorial31.html");
      if (workbook.easy_LoadHTMLFile(file))
      {
        // Set worksheet name
        workbook.easy_getSheetAt(0).setSheetName("First tab");

        // Add new worksheet and add some data in cells (optional step)
        workbook.easy_addWorksheet("Second tab");
        ExcelTable xlsTable = ((ExcelWorksheet)workbook.easy_getSheetAt(1)).easy_getExcelTable();
        xlsTable.easy_getCell("A1").setValue("Data added by Tutorial40");

        for (int column=0; column<5; column++)
        {
          xlsTable.easy_getCell(1, column).setValue("Data " + (column + 1));
        }

        // Export Excel file
        System.out.println("Writing file C:\\Samples\\Tutorial40 - convert HTML to Excel.xlsx.");
        workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial40 - convert HTML to Excel.xlsx");

        // Confirm conversion of HTML to Excel
        if (workbook.easy_getError().equals(""))
          System.out.println("File successfully created.");
        else
          System.out.println("Error encountered: " + workbook.easy_getError());

      }
	  
	  // Dispose memory
      workbook.Dispose();
    }
    catch (Exception ex) {
      ex.printStackTrace();
    }
  }
}