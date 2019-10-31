/* ----------------------------------------------------------------
 * Tutorial 13
 * 
 * This tutorial shows how to create a Microsoft Excel file
 * that has two worksheets. The second one contains a named
 * range. The first 10 rows of the first 2 columns contain
 * validators.
 * ----------------------------------------------------------------- */

using System;
using EasyXLS;
using EasyXLS.Constants;

public class Tutorial13
{

	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 13\n----------\n");

		//Create an instance of the object that generates Excel files, having 2 sheets
		ExcelDocument xls = new ExcelDocument(2);
	    
		//Set the sheet names
		xls.easy_getSheetAt(0).setSheetName("First tab");
		xls.easy_getSheetAt(1).setSheetName("Second tab");

		//Get the table of the second worksheet and populate the sheet
		ExcelWorksheet xlsSecondTab = (ExcelWorksheet)xls.easy_getSheetAt(1);
		ExcelTable xlsSecondTable = xlsSecondTab.easy_getExcelTable();
		xlsSecondTable.easy_getCell("A1").setValue("Range data 1");
		xlsSecondTable.easy_getCell("A2").setValue("Range data 2");
		xlsSecondTable.easy_getCell("A3").setValue("Range data 3");
		xlsSecondTable.easy_getCell("A4").setValue("Range data 4");

		//Create a named range
		xlsSecondTab.easy_addName("Range", "=Second tab!$A$1:$A$4");

		//Add a validator for the first 10 rows of the first column
		ExcelWorksheet xlsFirstTab = (ExcelWorksheet)xls.easy_getSheetAt(0);
		xlsFirstTab.easy_addDataValidator("A1:A10", DataValidator.VALIDATE_LIST, DataValidator.OPERATOR_EQUAL_TO, "=Range", "");

		//Add a validator for the first 10 rows of the second column
		xlsFirstTab.easy_addDataValidator("B1:B10", DataValidator.VALIDATE_WHOLE_NUMBER, DataValidator.OPERATOR_BETWEEN, "=4", "=100");

		//Generate the file
		Console.WriteLine("Writing file C:\\Samples\\Tutorial13.xls.");
		xls.easy_WriteXLSFile("C:\\Samples\\Tutorial13.xls");

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

