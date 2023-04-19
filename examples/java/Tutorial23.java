//package testexceljava;

import java.awt.Color;
import EasyXLS.*;
import EasyXLS.Charts.*;
import EasyXLS.Constants.*;
import EasyXLS.Drawings.Formatting.*;


/*-----------------------------------------------------------------
 | Tutorial 23
 | 
 | This tutorial shows how to create an Excel file in Java with a
 | chart and how to set chart type and formatting properties for
 | chart area, plot area, axis, series and legend.
  ---------------------------------------------------------------*/

public class Tutorial23 {

  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 23");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files
      ExcelDocument workbook = new ExcelDocument();

      // Create an worksheet
      workbook.easy_addWorksheet("SourceData");

      // Get the table of data for the worksheet
      ExcelTable xlsTable1 = ((ExcelWorksheet)workbook.easy_getSheet("SourceData")).easy_getExcelTable();

	  // Add data in cells for report header
      xlsTable1.easy_getCell(0, 0).setValue("Show Date");
      xlsTable1.easy_getCell(0, 1).setValue("Available Places");
      xlsTable1.easy_getCell(0, 2).setValue("Available Tickets");
      xlsTable1.easy_getCell(0, 3).setValue("Sold Tickets");

	  // Add data in cells for chart report values
      xlsTable1.easy_getCell(1, 0).setValue("03/13/2005 00:00:00");
      xlsTable1.easy_getCell(1, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE);
      xlsTable1.easy_getCell(2, 0).setValue("03/14/2005 00:00:00");
      xlsTable1.easy_getCell(2, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE);
      xlsTable1.easy_getCell(3, 0).setValue("03/15/2005 10:00:00");
      xlsTable1.easy_getCell(3, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE);
      xlsTable1.easy_getCell(4, 0).setValue("03/16/2005 00:00:00");
      xlsTable1.easy_getCell(4, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE);

      xlsTable1.easy_getCell(1, 1).setValue("10000");
      xlsTable1.easy_getCell(2, 1).setValue("5000");
      xlsTable1.easy_getCell(3, 1).setValue("8500");
      xlsTable1.easy_getCell(4, 1).setValue("1000");

      xlsTable1.easy_getCell(1, 2).setValue("8000");
      xlsTable1.easy_getCell(2, 2).setValue("4000");
      xlsTable1.easy_getCell(3, 2).setValue("6000");
      xlsTable1.easy_getCell(4, 2).setValue("1000");

      xlsTable1.easy_getCell(1, 3).setValue("920");
      xlsTable1.easy_getCell(2, 3).setValue("1005");
      xlsTable1.easy_getCell(3, 3).setValue("342");
      xlsTable1.easy_getCell(4, 3).setValue("967");

	  // Set column widths
      xlsTable1.easy_getColumnAt(0).setWidth(100);
      xlsTable1.easy_getColumnAt(1).setWidth(100);
      xlsTable1.easy_getColumnAt(2).setWidth(100);
      xlsTable1.easy_getColumnAt(3).setWidth(100);
      
      // Add a chart sheet
     workbook.easy_addChart("Chart", "=SourceData!$A$1:$D$5", Chart.SERIES_IN_COLUMNS);

     // Get the previously added chart
     ExcelChart xlsChart = ((ExcelChartSheet)workbook.easy_getSheetAt(1)).easy_getExcelChart();

      // Set chart type
      xlsChart.easy_setChartType(Chart.CHART_TYPE_CYLINDER_COLUMN);

      // Format chart area
      ExcelChartArea xlsChartArea = xlsChart.easy_getChartArea();
      xlsChartArea.getLineColorFormat().setLineColor(Color.darkGray);
      xlsChartArea.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_SOLID);
      xlsChartArea.getLineStyleFormat().setWidth(0.25f);

      // Format chart plot area
      ExcelPlotArea xlsPlotArea =  xlsChart.easy_getPlotArea();
      xlsPlotArea.getLineColorFormat().setLineColor(Color.darkGray);
      xlsPlotArea.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_SOLID);
      xlsPlotArea.getLineStyleFormat().setWidth(0.25f);

      // Format chart legend
      ExcelChartLegend xlsChatLegend = xlsChart.easy_getLegend();
      xlsChatLegend.getFillFormat().setBackground(Color.pink);
      xlsChatLegend.getFontFormat().setForeground(Color.blue);
      xlsChatLegend.getFontFormat().setItalic(true);
      xlsChatLegend.setKeysArrangementDirection(Chart.KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL);
      xlsChatLegend.setPlacement(Chart.LEGEND_CORNER);
      xlsChatLegend.getShadowFormat().setShadow(ShadowFormat.OFFSET_DIAGONAL_BOTTOM_RIGHT);

      // Format chart X axis
      ExcelAxis xlsXAxis = xlsChart.easy_getCategoryXAxis();
      xlsXAxis.getLineColorFormat().setLineColor(Color.lightGray);
      xlsXAxis.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_DASH_DOT);
      xlsXAxis.getLineStyleFormat().setWidth(0.25f);
      xlsXAxis.getFontFormat().setForeground(Color.red);

      // Format chart Y axis
      ExcelAxis xlsYAxis = xlsChart.easy_getValueYAxis();
      xlsYAxis.getLineColorFormat().setLineColor(Color.lightGray);
      xlsYAxis.getLineStyleFormat().setDashType(LineStyleFormat.DASH_TYPE_LONG_DASH);
      xlsYAxis.getLineStyleFormat().setWidth(0.25f);
      xlsYAxis.getFontFormat().setForeground(Color.blue);

      // Fomat chart series
      xlsChart.easy_getSeriesAt(0).getFillFormat().setBackground(Color.blue);
      xlsChart.easy_getSeriesAt(1).getFillFormat().setBackground(Color.yellow);
      xlsChart.easy_getSeriesAt(2).getFillFormat().setBackground(Color.green);

      // Export Excel file
      System.out.println("Writing file: C:\\Samples\\Tutorial23 - various Excel chart settings.xlsx");
      workbook.easy_WriteXLSXFile("C:\\Samples\\Tutorial23 - various Excel chart settings.xlsx");

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
