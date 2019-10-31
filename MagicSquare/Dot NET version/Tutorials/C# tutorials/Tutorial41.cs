/* ----------------------------------------------------------------
 * Tutorial 41
 * 
 * This tutorial shows how to load an XML file (we use the file
 * generated in Tutorial 32), modify some data and save it to
 * another file (Tutorial41.xls).
 * ----------------------------------------------------------------- */

using System;
using System.IO;
using System.Data;
using EasyXLS;

public class Tutorial41
{
	
	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 41\n-----------\n");

		//Create an instance of the object that generates Excel files
		ExcelDocument xls = new ExcelDocument();

		//Read the file
		Console.WriteLine("Reading file C:\\Samples\\Tutorial32.xml");
		
		if (xls.easy_LoadXMLSpreadsheetFile("C:\\Samples\\Tutorial32.xml"))
		{
			//Get the table of the second worksheet
			ExcelTable xlsSecondTable = ((ExcelWorksheet)xls.easy_getSheet("Second tab")).easy_getExcelTable();
			xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial41");
						
			for (int column=0; column<5; column++)
			{
				xlsSecondTable.easy_getCell(1, column).setValue("Data " + (column + 1)); 
			}
		
			//Generate the file
			Console.WriteLine("\nWriting file C:\\Samples\\Tutorial41.xls.");
			xls.easy_WriteXLSFile("C:\\Samples\\Tutorial41.xls");

			//Confirm generation
			String sError = xls.easy_getError();
			if (sError.Equals(""))
				Console.Write("\nFile successfully created.");
			else
				Console.Write("\nError encountered: " + sError);
		}
		else
		{
			Console.WriteLine("\nError reading file C:\\Samples\\Tutorial32.xml \n" + xls.easy_getError());	
		}

		Console.WriteLine("\nPress Enter to exit ...");
		
		//Dispose memory
		xls.Dispose();

		Console.ReadLine();
	}
}
