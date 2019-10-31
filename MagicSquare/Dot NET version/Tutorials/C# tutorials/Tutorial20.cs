/* ----------------------------------------------------------------
 * Tutorial 20
 * 
 * This tutorial shows how to create a Microsoft Excel file 
 * that has AutoFilter.
 * ----------------------------------------------------------------- */

using System;
using EasyXLS;
using EasyXLS.Constants;


public class Tutorial20
{
	
	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 20\n-----------\n");

		//Create an instance of the object that generates Excel files, having 1 sheet
		ExcelDocument xls = new ExcelDocument(1);
	 
		//Get the table of the first worksheet 
		ExcelWorksheet xlsTab = ((ExcelWorksheet)xls.easy_getSheet("Sheet1"));
		ExcelTable xlsTable = xlsTab.easy_getExcelTable();

		//Add the cells for header
		for (int column=0; column<5; column++)
		{
			xlsTable.easy_getCell(0,column).setValue("Column " + (column + 1)); 
			xlsTable.easy_getCell(0,column).setDataType(DataType.STRING);
		}
		
		//Add the cells for data
		for (int row=0; row<100; row++)
		{
			for (int column=0; column<5; column++)
			{
				xlsTable.easy_getCell(row+1,column).setValue("Data " + (row + 1) + ", " + (column + 1)); 
				xlsTable.easy_getCell(row+1,column).setDataType(DataType.STRING);
			}
		}
		
		//Add AutoFilter
		ExcelFilter xlsFilter = xlsTab.easy_getFilter();
		xlsFilter.setAutoFilter("A1:E1");
		
		//Generate the file
		Console.WriteLine("Writing file C:\\Samples\\Tutorial20.xls.");
		xls.easy_WriteXLSFile("C:\\Samples\\Tutorial20.xls");

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
