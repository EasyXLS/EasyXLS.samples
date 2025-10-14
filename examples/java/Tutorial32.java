//package testexceljava;

import EasyXLS.*;
import EasyXLS.Constants.*;

/*------------------------------------------------------------------------
 | Tutorial 32
 |
 | This tutorial shows how to export data to XML Spreadsheet file in Java.
 ------------------------------------------------------------------------*/

public class Tutorial32{

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 32");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files, having two sheets
      ExcelDocument workbook = new ExcelDocument(2);

      // Set the sheet names
      workbook.easy_getSheetAt(0).setSheetName("First tab");
      workbook.easy_getSheetAt(1).setSheetName("Second tab");

      // Get the table of data for the first worksheet
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

      // Apply a predefined format to the cells.
      xlsFirstTable.easy_setRangeAutoFormat("A1:E101",new ExcelAutoFormat(Styles.AUTOFORMAT_EASYXLS1));

      // Export XML Spreadsheet file
      System.out.println("Writing file: C:\\Samples\\Tutorial32 - export XML spreadsheet file.xml");
      workbook.easy_WriteXMLFile("C:\\Samples\\Tutorial32 - export XML spreadsheet file.xml");

      // Confirm export of XML file
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

