/* ----------------------------------------------------------------
 * Tutorial 37
 * 
 * This tutorial shows how to load a XLSX file (we use the file
 * generated in Tutorial 28), modify some data and save it to
 * another file (Tutorial37.xlsx).
 * ----------------------------------------------------------------- */

using System;
using System.IO;
using System.Data;
using EasyXLS;

public class Tutorial37
{

	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 37\n-----------\n");

		//Create an instance of the object that generates Excel files
		ExcelDocument xls = new ExcelDocument();

		//Read the file
		Console.WriteLine("Reading file C:\\Samples\\Tutorial28.xlsx");
		
		if (xls.easy_LoadXLSXFile("C:\\Samples\\Tutorial28.xlsx"))
		{
			//Get the table of the second worksheet
			ExcelTable xlsSecondTable = ((ExcelWorksheet)xls.easy_getSheet("Second tab")).easy_getExcelTable();

			//Write some data
			xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial37");
			
			for (int column=0; column<5; column++)
			{
				xlsSecondTable.easy_getCell(1, column).setValue("Data " + (column + 1)); 
			}
		
			//Generate the file
			Console.WriteLine("\nWriting file C:\\Samples\\Tutorial37.xlsx.");
			xls.easy_WriteXLSXFile("C:\\Samples\\Tutorial37.xlsx");

			//Confirm generation
			String sError = xls.easy_getError();
			if (sError.Equals(""))
				Console.Write("\nFile successfully created.");
			else
				Console.Write("\nError encountered: " + sError);
		}
		else
		{
			Console.WriteLine("\nError reading file C:\\Samples\\Tutorial28.xlsx \n" + xls.easy_getError());	
		}

		Console.WriteLine("\nPress Enter to exit ...");
		
		//Dispose memory
		xls.Dispose();

		Console.ReadLine();
	}
}
