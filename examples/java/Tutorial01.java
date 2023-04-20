//package testexceljava;

import java.sql.ResultSet;
import java.sql.*;
import EasyXLS.*;
import EasyXLS.Constants.*;

/*----------------------------------------------------------------------
 | Tutorial 01
 |
 | This code sample shows how to export ResultSet to Excel file in Java.
 | The ResultSet contains data from a SQL database.
 | The cells are formatted using a predefined format.
 ----------------------------------------------------------------------*/

public class Tutorial01 {


  public static void main(String[] args) {
    try {
      System.out.println("Tutorial 01");
      System.out.println("----------");

      // Create an instance of the class that exports Excel files
      ExcelDocument workbook = new ExcelDocument();

      // Create the database connection
      Class.forName("com.microsoft.jdbc.sqlserver.SQLServerDriver");
      String sConnectionString = "jdbc:microsoft:sqlserver://localhost:1433;databasename=Northwind;user=sa;password=;";
      Connection sqlConnection = (Connection) DriverManager.getConnection(sConnectionString);

      // Create the statement used to populate the resultset and populate the resultset
      String sQueryString = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', P.ProductName AS 'Product Name', O.UnitPrice AS Price, O.Quantity , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID";
      PreparedStatement pStatement = sqlConnection.prepareStatement(sQueryString, java.sql.ResultSet.TYPE_SCROLL_INSENSITIVE, java.sql.ResultSet.CONCUR_UPDATABLE);
      ResultSet rs = pStatement.executeQuery();

      // Export the Excel file
      System.out.println("Writing file C:\\Samples\\Tutorial01 - export ResultSet to Excel.xlsx.");
      workbook.easy_WriteXLSXFile_FromResultSet("c:\\Samples\\Tutorial01 - export ResultSet to Excel.xlsx", rs, new ExcelAutoFormat(Styles.AUTOFORMAT_EASYXLS1), "Sheet1");

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
