//package testexceljava;

/*------------------------------------------------------------------------------
 | Tutorial 36
 | 
 | This tutorial shows how to read an Excel XLSX file in Java (the
 | XLSX file generated by Tutorial 04 as base template), modify
 | some data and save it to another XLSX file (Tutorial36 - read XLSX file.xlsx).
 ------------------------------------------------------------------------------*/

import java.io.FileInputStream;
import EasyXLS.*;

public class Tutorial36 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 36");
      System.out.println("----------");

      // Create an instance of the class that reads Excel files
      ExcelDocument workbook = new ExcelDocument();

      // Read XLSX file
      System.out.println("Reading file C:\\Samples\\Tutorial04.xlsx.");
      FileInputStream file = new FileInputStream("C:\\Samples\\Tutorial04.xlsx");
      if (workbook.easy_LoadXLSXFile(file))
      {
		// Get the table of data for the second worksheet
        ExcelTable xlsSecondTable = ((ExcelWorksheet)workbook.easy_getSheetAt(1)).easy_getExcelTable();
		
		// Write some data to the second sheet
        xlsSecondTable .easy_getCell("A1").setValue("Data added by Tutorial36");

        for (int column=0; column<5; column++)
        {
          xlsSecondTable .easy_getCell(1, column).setValue("Data " + (column + 1));
        }

        // Export the new XLSX file
        System.out.println("Writing file C:\\Samples\\Tutorial36 - read XLSX file.xlsx.");
        workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial36 - read XLSX file.xlsx");

        // Confirm export of Excel file
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
