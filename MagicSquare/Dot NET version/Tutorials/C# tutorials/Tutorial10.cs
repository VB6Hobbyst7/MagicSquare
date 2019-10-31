/* -----------------------------------------------------------------
 * Tutorial 10
 * 
 * This tutorial shows how to merge a cell range.
 * ----------------------------------------------------------------- */

using System;
using EasyXLS;

public class Tutorial10
{

	[STAThread]
	static void Main()
	{
		Console.WriteLine("Tutorial 10\n-----------\n");

		//Create an instance of the object that generates Excel files
		ExcelDocument xls = new ExcelDocument(1);

		//Get the table of the first sheet
		ExcelTable xlsTable = ((ExcelWorksheet)xls.easy_getSheet("Sheet1")).easy_getExcelTable();

		//Merging cells
		xlsTable.easy_mergeCells("A1:C3");    
			
		//Generate the file
		Console.WriteLine("Writing file C:\\Samples\\Tutorial10.xls.");
		xls.easy_WriteXLSFile("C:\\Samples\\Tutorial10.xls");

		//Confirm generation
		String sError = xls.easy_getError();
		if (sError.Equals(""))
			Console.Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console.Write("\nError encountered: " + sError + "\nPress Enter to Exit...");
		
		//Dispose memory
		xls.Dispose();

		Console.ReadLine();
	}
}
