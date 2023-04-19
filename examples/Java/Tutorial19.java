//package testexceljava;

import EasyXLS.*;
import EasyXLS.Constants.*;

/*-------------------------------------------------------------------
| Tutorial 19
 |
 | This tutorial shows how to create an Excel file in Java having
 | multiple sheets. The first sheet is filled with data and the
 | afirst cell of the second row contains data in rich text format.
  -----------------------------------------------------------------*/

public class Tutorial19 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 19");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files having two sheets
      ExcelDocument workbook = new ExcelDocument(2);

      // Set the sheet names
      workbook.easy_getSheetAt(0).setSheetName("First tab");
      workbook.easy_getSheetAt(1).setSheetName("Second tab");

      // Get the table of data for the first worksheet
      ExcelTable xlsFirstTable = ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_getExcelTable();

      // Create the string used to set the RTF in cell
      String sFormattedValue = "This is <b>bold</b>.";
      sFormattedValue += "\nThis is <i>italic</i>.";
      sFormattedValue += "\nThis is <u>underline</u>.";
      sFormattedValue += "\nThis is <underline double>double underline</underline double>.";
      sFormattedValue += "\nThis is <font color=red>red</font>.";
      sFormattedValue += "\nThis is <font color=rgb(255,0,0)>red</font> too.";
      sFormattedValue += "\nThis is <font face=\"Arial Black\">Arial Black</font>.";
      sFormattedValue += "\nThis is <font size=15pt>size 15</font>.";
      sFormattedValue += "\nThis is <s>strikethrough</s>.";
      sFormattedValue += "\nThis is <sup>superscript</sup>.";
      sFormattedValue += "\nThis is <sub>subscript</sub>.";
      sFormattedValue += "\n<b>This</b> <i>is</i> <font color=red face=\"Arial Black\" size=15pt><underline double>formatted</underline double></font> <s>text</s>.";

      // Set the rich text value in cell
      xlsFirstTable.easy_getCell(1, 0).setHTMLValue(sFormattedValue);
      xlsFirstTable.easy_getCell(1, 0).setDataType(DataType.STRING);
      xlsFirstTable.easy_getCell(1, 0).setWrap(true);
      xlsFirstTable.easy_getRowAt(1).setHeight(250);
      xlsFirstTable.easy_getColumnAt(0).setWidth(250);

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial19 - RTF for Excel cells.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial19 - RTF for Excel cells.xlsx");

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
