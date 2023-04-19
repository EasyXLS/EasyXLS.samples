//package testexceljava;

import java.sql.ResultSet;
import java.sql.*;
import java.awt.Color;
import EasyXLS.*;
import EasyXLS.Constants.*;

/*------------------------------------------------------------------------
 | Tutorial 02
 |
 | This code sample shows how to export ResultSet to Excel file in Java.
 | The ResultSet contains data from a SQL database.
 | The cells are formatted using an user-defined format.
 -----------------------------------------------------------------------*/

public class Tutorial02 {


  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 02");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files
      ExcelDocument workbook = new ExcelDocument();

      // Create the database connection
      Class.forName("com.microsoft.jdbc.sqlserver.SQLServerDriver");
      String sConnectionString = "jdbc:microsoft:sqlserver://localhost:1433;databasename=Northwind;user=sa;password=;";
      Connection sqlConnection = (Connection) DriverManager.getConnection(sConnectionString);

      // Create the statement used to populate the resultset and populate the resultset
      String sQueryString = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', P.ProductName AS 'Product Name', O.UnitPrice AS Price, ' ' + cast(O.Quantity AS varchar) AS Quantity  , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID";
      PreparedStatement pStatement = sqlConnection.prepareStatement(sQueryString, java.sql.ResultSet.TYPE_SCROLL_INSENSITIVE, java.sql.ResultSet.CONCUR_UPDATABLE);
      ResultSet rs = pStatement.executeQuery();

      // Create an instance of the class used to format the cells in the report
      ExcelAutoFormat xlsAutoFormat = new ExcelAutoFormat();
      
	  // Set the formatting style of the header
      ExcelStyle xlsHeaderStyle = new ExcelStyle(new Color(144, 238, 144));
      xlsHeaderStyle.setFontSize(12);
      xlsAutoFormat.setHeaderRowStyle(xlsHeaderStyle);

      // Set the formatting style of the cells (alternating style)
      ExcelStyle xlsEvenRowStripesStyle = new ExcelStyle(new Color(255, 250, 240));
      xlsEvenRowStripesStyle.setFormat("$0.00");
      xlsEvenRowStripesStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT);
      xlsAutoFormat.setEvenRowStripesStyle(xlsEvenRowStripesStyle);
      ExcelStyle xlsOddRowStripesStyle = new ExcelStyle(new Color(240, 247, 239));
      xlsOddRowStripesStyle.setFormat("$0.00");
      xlsOddRowStripesStyle.setHorizontalAlignment (Alignment.ALIGNMENT_LEFT);
      xlsAutoFormat.setOddRowStripesStyle(xlsOddRowStripesStyle);
      ExcelStyle xlsLeftColumnStyle = new ExcelStyle(new Color(255, 250, 240));
      xlsLeftColumnStyle.setFormat("mm/dd/yyyy");
      xlsLeftColumnStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT);
      xlsAutoFormat.setLeftColumnStyle(xlsLeftColumnStyle);

      // Export the Excel file
      System.out.println("Writing file C:\\Samples\\Tutorial02 - export ResultSet to Excel with formatting.xlsx.");
      workbook.easy_WriteXLSXFile_FromResultSet("c:\\Samples\\Tutorial02 - export ResultSet to Excel with formatting.xlsx", rs, xlsAutoFormat, "Sheet1");

      // Confirm export of Excel file
      if (workbook.easy_getError().equals(""))
        System.out.println("File successfully created.");
      else
        System.out.println("Error encountered: " + workbook.easy_getError());

      // Close the database connection
      sqlConnection.close();

      // Dispose memory
      workbook.Dispose();
    }
    catch (Exception ex) {
      ex.printStackTrace();
    }
  }
}
