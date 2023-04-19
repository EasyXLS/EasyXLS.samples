//package testexceljava;

import java.io.FileInputStream;
import java.sql.ResultSet;
import EasyXLS.*;

/*------------------------------------------------------------------------
 | Tutorial 34
 | 
 | This tutorial shows how to import Excel to ResultSet in Java. The data 
 | is imported from the active sheet of the Excel file (the Excel file 
 | generated in Tutorial 09).
 -----------------------------------------------------------------------*/

public class Tutorial34 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 34");
      System.out.println("----------");

      // Create an instance of the class that imports Excel files
      ExcelDocument workbook = new ExcelDocument();

      // Import Excel file to ResultSet
      System.out.println("Reading file C:\\Samples\\Tutorial09.xlsx.");
      FileInputStream file = new FileInputStream("C:\\Samples\\Tutorial09.xlsx");
      ResultSet rs = workbook.easy_ReadXLSXActiveSheet_AsResultSet(file);
      
	  // Display imported ResultSet values
      int columnCount = rs.getMetaData().getColumnCount();
      int row = 0;
      while (rs.next()){
        for (int column=1; column<=columnCount; column++)
          System.out.println("At row " + (row + 1) + ", column " + (column) +
                             " the value is '" + rs.getString(column) + "'");
        row++;
      }

      // Dispose memory
      workbook.Dispose();
    }
    catch (Exception ex) {
      ex.printStackTrace();
    }
  }
}
