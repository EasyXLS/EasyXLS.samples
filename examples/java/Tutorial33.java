//package testexceljava;

import EasyXLS.*;
import EasyXLS.Constants.*;

/*----------------------------------------------------------------------------
 | Tutorial 33
 |
 | This tutorial shows how to set document properties for Excel file in Java,
 | like 'Subject' property for summary information, 'Manager' property for
 | document summary information and a custom property.
  --------------------------------------------------------------------------*/

public class Tutorial33 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 33");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files
      ExcelDocument workbook = new ExcelDocument(1);

      // Set the 'Subject' document property
      workbook.getSummaryInformation().setSubject("This is the subject");

      // Set the 'Manager' document property
      workbook.getDocumentSummaryInformation().setManager("This is the manager");

      // Set a custom document property
      workbook.getDocumentSummaryInformation().setCustomProperty("PropertyName", FileProperty.VT_NUMBER, "4");

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial33 - Excel file properties.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial33 - Excel file properties.xlsx");

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
