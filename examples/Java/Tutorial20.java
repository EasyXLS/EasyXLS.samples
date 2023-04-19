//package testexceljava;

import EasyXLS.*;
import EasyXLS.Constants.*;

/*---------------------------------------------------------
 | Tutorial 20
 |
 | This tutorial shows how to create an Excel file in Java
 | and apply an auto-filter to a range of cells.
  -------------------------------------------------------*/

public class Tutorial20 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 20");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files having one sheet
      ExcelDocument workbook = new ExcelDocument(1);

      // Get the table of data for the worksheet
      ExcelWorksheet xlsTab = ((ExcelWorksheet)workbook.easy_getSheet("Sheet1"));
      ExcelTable xlsTable = xlsTab.easy_getExcelTable();

      // Add data in cells for report header
      for (int column=0; column<5; column++)
      {
        xlsTable.easy_getCell(0,column).setValue("Column " + (column + 1));
        xlsTable.easy_getCell(0,column).setDataType(DataType.STRING);
      }

      // Add data in cells for report values
      for (int row=0; row<100; row++)
      {
        for (int column=0; column<5; column++)
        {
          xlsTable.easy_getCell(row+1,column).setValue("Data " + (row + 1) + ", " + (column + 1));
          xlsTable.easy_getCell(row+1,column).setDataType(DataType.STRING);
        }
      }

      // Apply auto-filter on cell range A1:E1
      ExcelFilter xlsFilter = xlsTab.easy_getFilter();
      xlsFilter.setAutoFilter("A1:E1");

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial20 - autofilter in Excel sheet.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial20 - autofilter in Excel sheet.xlsx");

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
