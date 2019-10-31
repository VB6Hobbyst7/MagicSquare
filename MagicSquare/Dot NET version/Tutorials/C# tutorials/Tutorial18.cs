/* ----------------------------------------------------------------
 * Tutorial 18
 * 
 * This tutorial shows how to create a Microsoft Excel file
 * that has two worksheets. The first one is full with data 
 * and the panes are frozen.
 * ----------------------------------------------------------------- */

using System;
using EasyXLS;
using EasyXLS.Constants;

public class Tutorial18
{
	
	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 18\n----------\n");

		//Create an instance of the object that generates Excel files, having 2 sheets
		ExcelDocument xls = new ExcelDocument(2);
	    
		//Set the sheet names
		xls.easy_getSheetAt(0).setSheetName("First tab");
		xls.easy_getSheetAt(1).setSheetName("Second tab");

		//Get the table of the first worksheet
		ExcelTable xlsFirstTable = ((ExcelWorksheet)xls.easy_getSheetAt(0)).easy_getExcelTable();

		//Add the cells for header
		for (int column=0; column<5; column++)
		{
			xlsFirstTable.easy_getCell(0,column).setValue("Column " + (column + 1)); 
			xlsFirstTable.easy_getCell(0,column).setDataType(DataType.STRING);
		}
		xlsFirstTable.easy_getRowAt(0).setHeight(30);

		//Add the cells for data
		for (int row=0; row<100; row++)
		{
			for (int column=0; column<5; column++)
			{
				xlsFirstTable.easy_getCell(row+1,column).setValue("Data " + (row + 1) + ", " + (column + 1)); 
				xlsFirstTable.easy_getCell(row+1,column).setDataType(DataType.STRING);
			}
		}

		//Set column widths
		xlsFirstTable.setColumnWidth(0, 70);
		xlsFirstTable.setColumnWidth(1, 100);
		xlsFirstTable.setColumnWidth(2, 70);
		xlsFirstTable.setColumnWidth(3, 100);
		xlsFirstTable.setColumnWidth(4, 70);

		//Freeze panes
		xlsFirstTable.easy_freezePanes(1, 0, 75, 0);


		//Generate the file
		Console.WriteLine("Writing file C:\\Samples\\Tutorial18.xls.");
		xls.easy_WriteXLSFile("C:\\Samples\\Tutorial18.xls");

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

