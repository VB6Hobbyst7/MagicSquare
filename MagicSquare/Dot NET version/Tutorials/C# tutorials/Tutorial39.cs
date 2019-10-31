/* ----------------------------------------------------------------
 * Tutorial 39
 * 
 * This tutorial shows how to load a CSV file (we use the file
 * generated in Tutorial 30), modify some data and save it to
 * another file (Tutorial39.xls).
 * ----------------------------------------------------------------- */

using System;
using System.IO;
using System.Data;
using EasyXLS;

public class Tutorial39
{

	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 39\n-----------\n");

		//Create an instance of the object that generates Excel files
		ExcelDocument xls = new ExcelDocument();

		//Read the file
		Console.WriteLine("Reading file C:\\Samples\\Tutorial30.csv");
		
		if (xls.easy_LoadCSVFile("C:\\Samples\\Tutorial30.csv"))
		{
			//Set the name of the first worksheet
			xls.easy_getSheetAt(0).setSheetName("First tab");

			//Add a new worksheet and write some data
			xls.easy_addWorksheet("Second tab");
			ExcelTable xlsTable = ((ExcelWorksheet)xls.easy_getSheetAt(1)).easy_getExcelTable();
			xlsTable.easy_getCell("A1").setValue("Data added by Tutorial39");
						
			for (int column=0; column<5; column++)
			{
				xlsTable.easy_getCell(1, column).setValue("Data " + (column + 1)); 
			}
		
			//Generate the file
			Console.WriteLine("\nWriting file C:\\Samples\\Tutorial39.xls.");
			xls.easy_WriteXLSFile("C:\\Samples\\Tutorial39.xls");

			//Confirm generation
			String sError = xls.easy_getError();
			if (sError.Equals(""))
				Console.Write("\nFile successfully created.");
			else
				Console.Write("\nError encountered: " + sError);
		}
		else
		{
			Console.WriteLine("\nError reading file C:\\Samples\\Tutorial30.csv \n" + xls.easy_getError());	
		}

		Console.WriteLine("\nPress Enter to exit ...");
		
		//Dispose memory
		xls.Dispose();

		Console.ReadLine();
	}
}
