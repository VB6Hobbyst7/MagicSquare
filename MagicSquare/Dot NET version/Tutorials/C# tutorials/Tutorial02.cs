/* ----------------------------------------------------------------
 * Tutorial 02
 * 
 * This tutorial shows how to generate an Excel document from a list of values. 
 * The cells are formatted using an user-defined format.
 * ----------------------------------------------------------------- */

using System;
using System.IO;
using System.Data;
using System.Drawing;
using EasyXLS;
using EasyXLS.Constants;



public class Tutorial02
{

	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 02\n-----------\n");

		// Create an instance of the object that generates Excel files
		ExcelDocument xls = new ExcelDocument();

		// Create the database connection
		String sConnectionString = "Initial Catalog=Northwind;Data Source=localhost;User ID=sa;Password=;";
		System.Data.SqlClient.SqlConnection sqlConnection = new System.Data.SqlClient.SqlConnection(sConnectionString);
		sqlConnection.Open();       		

		// Create the adapter used to fill the dataset
		String sQueryString = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', " +
			" P.ProductName AS 'Product Name', O.UnitPrice AS Price, cast(O.Quantity AS varchar) AS Quantity , O.UnitPrice * O. Quantity AS Value" +
			" FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID";
		System.Data.SqlClient.SqlDataAdapter adp = new System.Data.SqlClient.SqlDataAdapter(sQueryString, sqlConnection);

		// Populate the dataset
		DataSet ds  = new DataSet();
		adp.Fill(ds);


		// Create an instance of the object used to format the cells.
		ExcelAutoFormat xlsAutoFormat = new ExcelAutoFormat();
        // Set the style of the header
		ExcelStyle xlsHeaderStyle = new ExcelStyle(Color.LightGreen);
		xlsHeaderStyle.setFontSize(12);
		xlsAutoFormat.setHeaderRowStyle(xlsHeaderStyle);

		// Set the style of the cells
		ExcelStyle xlsEvenRowStripesStyle = new ExcelStyle(Color.FloralWhite);
		xlsEvenRowStripesStyle.setFormat("$0.00");
		xlsEvenRowStripesStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT);
		xlsAutoFormat.setEvenRowStripesStyle(xlsEvenRowStripesStyle);
		ExcelStyle xlsOddRowStripesStyle = new ExcelStyle(Color.FromArgb(240, 247, 239));
		xlsOddRowStripesStyle.setFormat("$0.00");
		xlsOddRowStripesStyle.setHorizontalAlignment (Alignment.ALIGNMENT_LEFT);
		xlsAutoFormat.setOddRowStripesStyle(xlsOddRowStripesStyle);
		ExcelStyle xlsLeftColumnStyle = new ExcelStyle(Color.FloralWhite);
		xlsLeftColumnStyle.setFormat("mm/dd/yyyy");
		xlsLeftColumnStyle.setHorizontalAlignment (Alignment.ALIGNMENT_LEFT);
		xlsAutoFormat.setLeftColumnStyle(xlsLeftColumnStyle);

		// Generate the file
		Console.WriteLine("Writing file C:\\Samples\\Tutorial02.xls.");
		xls.easy_WriteXLSFile_FromDataSet("c:\\Samples\\Tutorial02.xls", ds, xlsAutoFormat, "Sheet1");

		// Confirm generation
		String sError = xls.easy_getError();
		if (sError.Equals(""))
			Console.Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console.Write("\nError encountered: " + sError + "\nPress Enter to Exit...");


		// Close the database connection.
		sqlConnection.Close();

		// Dispose memory
		xls.Dispose();
		ds.Dispose();
		sqlConnection.Dispose();
		adp.Dispose();

		Console.ReadLine();
	}
}
