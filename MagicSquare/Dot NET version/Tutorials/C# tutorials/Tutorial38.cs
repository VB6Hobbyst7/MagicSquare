/* ----------------------------------------------------------------
 * Tutorial 38
 * 
 * This tutorial shows how to load a XLSB file (we use the file
 * generated in Tutorial 29), modify some data and save it to
 * another file (Tutorial38.xlsb).
 * ----------------------------------------------------------------- */

using System;
using System.IO;
using System.Data;
using EasyXLS;

public class Tutorial38
{

	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 38\n-----------\n");

		//Create an instance of the object that generates Excel files
		ExcelDocument xls = new ExcelDocument();

		//Read the file
		Console.WriteLine("Reading file C:\\Samples\\Tutorial29.xlsb");
		
		if (xls.easy_LoadXLSBFile("C:\\Samples\\Tutorial29.xlsb"))
		{
			//Get the table of the second worksheet
			ExcelTable xlsSecondTable = ((ExcelWorksheet)xls.easy_getSheet("Second tab")).easy_getExcelTable();

			//Write some data
			xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial38");
			
			for (int column=0; column<5; column++)
			{
				xlsSecondTable.easy_getCell(1, column).setValue("Data " + (column + 1)); 
			}
		
			//Generate the file
			Console.WriteLine("\nWriting file C:\\Samples\\Tutorial38.xlsb.");
			xls.easy_WriteXLSBFile("C:\\Samples\\Tutorial38.xlsb");

			//Confirm generation
			String sError = xls.easy_getError();
			if (sError.Equals(""))
				Console.Write("\nFile successfully created.");
			else
				Console.Write("\nError encountered: " + sError);
		}
		else
		{
			Console.WriteLine("\nError reading file C:\\Samples\\Tutorial29.xlsb \n" + xls.easy_getError());	
		}

		Console.WriteLine("\nPress Enter to exit ...");
		
		//Dispose memory
		xls.Dispose();

		Console.ReadLine();
	}
}
