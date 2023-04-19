//package testexceljava;

import java.awt.Color;
import EasyXLS.*;
import EasyXLS.Constants.*;

/*----------------------------------------------------------------------------
 | Tutorial 09
 |
 | This tutorial shows how to create an Excel file in Java 
 | having multiple sheets. The first sheet is filled with data
 | and the cells are formatted and locked.
 | The column header has comments.
 | The first worksheet has header & footer.
 | The first worksheet has print area, rows to repeat at top, center on page,
 | page orientation, page order, paper size, comments print location,
 | print gridlines option and page breaks.
 ---------------------------------------------------------------------------*/

public class Tutorial09 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 09");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files, having two sheets
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

        // Add comment for report header cells
        xlsFirstTable.easy_getCell(0, column).setComment("This is column no " + (column + 1));
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

      // Add header on center section 
      ExcelWorksheet xlsFirstTab = (ExcelWorksheet)workbook.easy_getSheetAt(0);
      xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertSingleUnderline();
      xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertFile();
      xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertValue(" - How to create header and footer");

	  // Add header on right section
      xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertDate();
      xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertValue(" ");
      xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertTime();

      // Add footer on center section
      xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertPage();
      xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertValue(" of ");
      xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertPages();

       // Get the object that stores the page setup options for the first sheet
      ExcelPageSetup xlsPageSetup = xlsFirstTab.easy_getPageSetup();
      // Set print area
	  xlsPageSetup.easy_setPrintArea("A1:E101");
      // Set the rows to repeat at top
	  xlsPageSetup.easy_setRowsToRepeatAtTop("$1:$1");
      // Set center on page option
	  xlsPageSetup.setCenterHorizontally(true);
      // Set page orientation
	  xlsPageSetup.setOrientation(PageSetup.ORIENTATION_PORTRAIT);
      // Set page order
	  xlsPageSetup.setPageOrder(PageSetup.PAGE_ORDER_DOWN_THEN_OVER);
      // Set paper size
	  xlsPageSetup.setPaperSize(PageSetup.PAPER_SIZE_A4);
      // Set where the comments to be printed
	  xlsPageSetup.setPrintComments(PageSetup.COMMENTS_AT_END_OF_SHEET);
      // Set the gridlines to be printed
	  xlsPageSetup.setPrintGridlines(true);
      
	  // Insert page breaks on rows
	  xlsFirstTable.easy_insertPageBreakAtRow(21);
      xlsFirstTable.easy_insertPageBreakAtRow(41);
      xlsFirstTable.easy_insertPageBreakAtRow(61);
      xlsFirstTable.easy_insertPageBreakAtRow(81);
      
	  // Set page break preview for the sheet
	  xlsFirstTab.setPageBreakPreview(true);

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial09 - Excel page setup.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial09 - Excel page setup.xlsx");

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
