//package testexceljava;

import java.awt.Color;
import EasyXLS.*;
import EasyXLS.Constants.*;

/*-----------------------------------------------------------
 | Tutorial 14
 | 
 | This tutorial shows how to create an Excel file in Java
 | having a sheet and conditional formatting for cell ranges.
  ----------------------------------------------------------*/

public class Tutorial14 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 14");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files having one sheet
      ExcelDocument workbook = new ExcelDocument(1);

      // Get the table of data for the first worksheet
      ExcelWorksheet xlsTab = ((ExcelWorksheet)workbook.easy_getSheet("Sheet1"));
      ExcelTable xlsTable = xlsTab.easy_getExcelTable();

      // Add data in cells
      for (int i=0;i<6;i++)
      {
        for (int j=0;j<4;j++)
        {
          if((i<2)&&(j<2))
            xlsTable.easy_getCell(i, j).setValue("12");
          else
            if((j==2)&&(i<2))
              xlsTable.easy_getCell(i, j).setValue("1000");
          else
            xlsTable.easy_getCell(i, j).setValue("9");
          xlsTable.easy_getCell(i, j).setDataType(DataType.NUMERIC);
        }
      }

      // Set conditional formatting
      xlsTab.easy_addConditionalFormatting("A1:C3", ConditionalFormatting.OPERATOR_BETWEEN, "=9", "=11", true, true, Color.RED);

      // Set another conditional formatting
      xlsTab.easy_addConditionalFormatting("A6:C6", ConditionalFormatting.OPERATOR_BETWEEN, "=COS(PI())+2", "", Color.ORANGE);
      xlsTab.easy_getConditionalFormattingAt("A6:C6").getConditionAt(0).setConditionType(ConditionalFormatting.CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA);

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial14 - conditional formatting in Excel.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial14 - conditional formatting in Excel.xlsx");

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
