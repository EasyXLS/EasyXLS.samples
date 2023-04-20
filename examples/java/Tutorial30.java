//package testexceljava;

import EasyXLS.*;
import EasyXLS.Constants.*;

/*-------------------------------------------------------------
 | Tutorial 30
 |
 | This tutorial shows how to export data to CSV file in Java.
 -------------------------------------------------------------*/

public class Tutorial30{

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 30");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files, having a sheet
      ExcelDocument workbook = new ExcelDocument(2);

      // Set the sheet name
      workbook.easy_getSheetAt(0).setSheetName("First tab");

      // Get the table of data for the worksheet
      ExcelTable xlsFirstTable =  ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_getExcelTable();

      // Add data in cells for report header
      for (int column=0; column<5; column++)
      {
        xlsFirstTable.easy_getCell(0,column).setValue("Column " + (column + 1));
        xlsFirstTable.easy_getCell(0,column).setDataType(DataType.STRING);
      }

      // Add data in cells for report values
      for (int row=0; row<100; row++)
      {
        for (int column=0; column<5; column++)
        {
          xlsFirstTable.easy_getCell(row+1,column).setValue("Data " + (row + 1) + ", " + (column + 1));
          xlsFirstTable.easy_getCell(row+1,column).setDataType(DataType.STRING);
        }
      }

      // Export CSV file
      System.out.println("Writing file: C:\\Samples\\Tutorial30 - export CSV file.csv");
      workbook.easy_WriteCSVFile("C:\\Samples\\Tutorial30 - export CSV file.csv", "First tab");

      // Confirm export of CSV file
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
