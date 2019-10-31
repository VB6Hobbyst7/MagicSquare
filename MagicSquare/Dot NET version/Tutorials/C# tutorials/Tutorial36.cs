/* ----------------------------------------------------------------
 * Tutorial 36
 * 
 * This tutorial shows how to load an excel file (we use the one
 * generated in Tutorial 09), modify some data and save it to
 * another file (Tutorial36.xls).
 * ----------------------------------------------------------------- */

using System;
using System.IO;
using System.Data;
using EasyXLS;

public class Tutorial36
{
	
	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 36\n-----------\n");

		//Create an instance of the object that generates Excel files
		ExcelDocument xls = new ExcelDocument();

		//Read the file
		Console.WriteLine("Reading file C:\\Samples\\Tutorial09.xls");
		
		if (xls.easy_LoadXLSFile("C:\\Samples\\Tutorial09.xls"))
		{
			//Get the table of the second worksheet
			ExcelTable xlsSecondTable = ((ExcelWorksheet)xls.easy_getSheet("Second tab")).easy_getExcelTable();

			//Write some data
			xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial36");
			
			for (int column=0; column<5; column++)
			{
				xlsSecondTable.easy_getCell(1, column).setValue("Data " + (column + 1)); 
			}
		
			//Generate the file
			Console.WriteLine("\nWriting file C:\\Samples\\Tutorial36.xls.");
			xls.easy_WriteXLSFile("C:\\Samples\\Tutorial36.xls");

			//Confirm generation
			String sError = xls.easy_getError();
			if (sError.Equals(""))
				Console.Write("\nFile successfully created.");
			else
				Console.Write("\nError encountered: " + sError);
		}
		else
		{
			Console.WriteLine("Error reading file C:\\Samples\\Tutorial09.xls");	
		}

		Console.WriteLine("\nPress Enter to exit ...");
		
		//Dispose memory
		xls.Dispose();

		Console.ReadLine();
	}
}
