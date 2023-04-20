//package testexceljava;

import java.awt.Color;
import EasyXLS.*;
import EasyXLS.Constants.*;

/*-----------------------------------------------------------------
 | Tutorial 06
 |
 | This code sample shows how to create an Excel file in Java with
 | multiple sheets. The first sheet is protected and 
 | filled with data. The cells are formatted and locked.
 -----------------------------------------------------------------*/

public class Tutorial06 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 06");
      System.out.println("----------");

      // Create an instance of the class that creates Excel files, having two sheets
      ExcelDocument workbook = new ExcelDocument(2);

      // Set the sheet names
      workbook.easy_getSheetAt(0).setSheetName("First tab");
      workbook.easy_getSheetAt(1).setSheetName("Second tab");

      // Protect first sheet
      workbook.easy_getSheetAt(0).setSheetProtected(true);

      // Get the table of data for the first worksheet
      ExcelTable xlsFirstTable =  ((ExcelWorksheet)workbook.easy_getSheetAt(0)).easy_getExcelTable();

      // Create the formatting style for the header
      ExcelStyle xlsStyleHeader = new ExcelStyle("Verdana", 8, true, true, Color.YELLOW);
      xlsStyleHeader.setBackground(Color.BLACK);
      xlsStyleHeader.setBorderColors(Color.GRAY, Color.GRAY, Color.GRAY, Color.GRAY);
      xlsStyleHeader.setBorderStyles(Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM);
      xlsStyleHeader.setHorizontalAlignment(Alignment.ALIGNMENT_CENTER);
      xlsStyleHeader.setVerticalAlignment(Alignment.ALIGNMENT_BOTTOM);
      xlsStyleHeader.setWrap(true);
      xlsStyleHeader.setDataType(DataType.STRING);

      // Add data in cells for report header
      for (int column=0; column<5; column++)
      {
        xlsFirstTable.easy_getCell(0, column).setValue("Column " + (column + 1));
        xlsFirstTable.easy_getCell(0, column).setStyle(xlsStyleHeader);
      }
      xlsFirstTable.easy_getRowAt(0).setHeight(30);

      // Add data in cells for report values
      for (int row=0; row<100; row++)
      {
        for (int column=0; column<5; column++)
        {
          xlsFirstTable.easy_getCell(row+1, column).setValue("Data " + (row + 1) + ", " + (column + 1));
        }
      }

      // Create a formatting style for cells
      ExcelStyle xlsStyleData = new ExcelStyle();
      xlsStyleData.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT);
      xlsStyleData.setForeground(Color.LIGHT_GRAY);
      xlsStyleData.setWrap(false);
      xlsStyleData.setDataType(DataType.STRING);
      // Protect cells
	  xlsStyleData.setLocked(true);
      xlsFirstTable.easy_setRangeStyle("A2:E101", xlsStyleData);

      // Set column widths
      xlsFirstTable.setColumnWidth(0, 70);
      xlsFirstTable.setColumnWidth(1, 100);
      xlsFirstTable.setColumnWidth(2, 70);
      xlsFirstTable.setColumnWidth(3, 100);
      xlsFirstTable.setColumnWidth(4, 70);

      // Create Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial06 - protect Excel sheet.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial06 - protect Excel sheet.xlsx");

      // Confirm the creation of Excel file
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
