/* ----------------------------------------------------------------
 * Tutorial 35
 * 
 * This tutorial shows how to read values from a sheet
 * of an excel file (For this example we use the file generated
 * in Tutorial 09).
 * ----------------------------------------------------------------- */

using System;
using System.IO;
using System.Data;
using EasyXLS;

class Tutorial35
{
	
	[STAThread]
	static void Main()
	{
		Console.WriteLine("Tutorial 35\n-----------\n");

		//Create an instance of the object that generates Excel files
		ExcelDocument xls = new ExcelDocument();

		//Read the file
		Console.WriteLine("Reading file C:\\Samples\\Tutorial09.xls.\n");
		DataSet ds = xls.easy_ReadXLSSheet_AsDataSet("C:\\Samples\\Tutorial09.xls", "First tab");

		//Display the values
		DataTable dt = ds.Tables[0];
		for (int row=0; row<dt.Rows.Count; row++)
			for (int column=0; column<dt.Columns.Count; column++)
				Console.WriteLine("At row " + (row + 1) + ", column " + (column + 1) +
					" the value is '" + dt.Rows[row].ItemArray[column] + "'");
 
		Console.Write("\nPress Enter to continue...");
		
		//Dispose memory
		xls.Dispose();

		Console.ReadLine();
	}
}

