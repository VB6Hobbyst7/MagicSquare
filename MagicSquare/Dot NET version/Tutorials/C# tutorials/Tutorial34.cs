/* ----------------------------------------------------------------
 * Tutorial 34
 * 
 * This tutorial shows how to read values from the active sheet
 * of an excel file (the file generated in Tutorial 09).
 * ----------------------------------------------------------------- */

using System;
using System.IO;
using System.Data;
using EasyXLS;

public class Tutorial34
{

	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 34\n-----------\n");

		//Create an instance of the object that reads Excel files
		ExcelDocument xls = new ExcelDocument();

		//Read the file
		Console.WriteLine("Reading file C:\\Samples\\Tutorial09.xls.\n");
		DataSet ds = xls.easy_ReadXLSActiveSheet_AsDataSet("C:\\Samples\\Tutorial09.xls");		

		//Display the values
		DataTable dt = ds.Tables[0];
		for (int row=0; row<dt.Rows.Count; row++)
			for (int column=0; column<dt.Columns.Count; column++)
				Console.WriteLine("At row " + (row + 1) + ", column " + (column + 1) +
					" the value is '" + dt.Rows[row].ItemArray[column] + "'");

		Console.Write("\nPress Enter to Exit...");
		
		//Dispose memory
		xls.Dispose();

		Console.ReadLine();
	}
}
