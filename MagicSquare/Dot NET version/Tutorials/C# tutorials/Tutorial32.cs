/* ----------------------------------------------------------------
 * Tutorial 32
 * 
 * This tutorial shows how to export an XML file.
 * ----------------------------------------------------------------- */

using System;
using EasyXLS;
using EasyXLS.Constants;

public class Tutorial32
{
	
	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 32\n----------\n");

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

		//Add the cells for data
		for (int row=0; row<100; row++)
		{
			for (int column=0; column<5; column++)
			{
				xlsFirstTable.easy_getCell(row+1,column).setValue("Data " + (row + 1) + ", " + (column + 1)); 
				xlsFirstTable.easy_getCell(row+1,column).setDataType(DataType.STRING);
			}
		}

		// Apply a predefined format to the cells.
		xlsFirstTable.easy_setRangeAutoFormat("A1:E101",new ExcelAutoFormat(Styles.AUTOFORMAT_EASYXLS1));

		//Generate the file
		Console.WriteLine("Writing file C:\\Samples\\Tutorial32.xml.");
		xls.easy_WriteXMLFile("C:\\Samples\\Tutorial32.xml");

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

