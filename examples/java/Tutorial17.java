//package testexceljava;

import EasyXLS.*;
import EasyXLS.Constants.*;

/*--------------------------------------------------------------------------------
 | Tutorial 17
 |
 | This tutorial shows how to create an Excel file with groups on rows in Java.
 | The Excel file has two worksheets. The first one is full with data and contains
 | the data groups.
  -------------------------------------------------------------------------------*/

public class Tutorial17 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 17");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files having two sheets
      ExcelDocument workbook = new ExcelDocument(2);

      // Set the sheet names
      workbook.easy_getSheetAt(0).setSheetName("First tab");
      workbook.easy_getSheetAt(1).setSheetName("Second tab");

      // Get the table of data for the first worksheet
      ExcelTable xlsFirstTable = ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_getExcelTable();

      // Add data in cells for report header
      for (int column=0; column<5; column++)
      {
        xlsFirstTable.easy_getCell(0,column).setValue("Column " + (column + 1));
        xlsFirstTable.easy_getCell(0,column).setDataType(DataType.STRING);
      }
      xlsFirstTable.easy_getRowAt(0).setHeight(30);

      // Add data in cells for report values
      for (int row=0; row<25; row++)
      {
        for (int column=0; column<5; column++)
        {
          xlsFirstTable.easy_getCell(row+1,column).setValue("Data " + (row + 1) + ", " + (column + 1));
          xlsFirstTable.easy_getCell(row+1,column).setDataType(DataType.STRING);
        }
      }

      // Set column widths
      xlsFirstTable.setColumnWidth(0, 70);
      xlsFirstTable.setColumnWidth(1, 100);
      xlsFirstTable.setColumnWidth(2, 70);
      xlsFirstTable.setColumnWidth(3, 100);
      xlsFirstTable.setColumnWidth(4, 70);

      // Group rows and format A1:E26 cell range
      ExcelDataGroup xlsFirstDataGroup = new ExcelDataGroup("A1:E26", DataGroup.GROUP_BY_ROWS, false);
      xlsFirstDataGroup .setAutoFormat(new ExcelAutoFormat(EasyXLS.Constants.Styles.AUTOFORMAT_EASYXLS1));
      ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_addDataGroup(xlsFirstDataGroup );

      // Group rows and format A2:E10 cell range, outline level two, inside previous group
      ExcelDataGroup xlsSecondDataGroup = new ExcelDataGroup("A2:E10", DataGroup.GROUP_BY_ROWS, false);
      xlsSecondDataGroup.setAutoFormat(new ExcelAutoFormat(EasyXLS.Constants.Styles.AUTOFORMAT_EASYXLS2));
      ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_addDataGroup(xlsSecondDataGroup);

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial17 - group data in Excel.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial17 - group data in Excel.xlsx");

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
