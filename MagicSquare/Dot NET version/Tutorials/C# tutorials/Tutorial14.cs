/* -----------------------------------------------------------------
 * Tutorial 14
 * 
 * This tutorial shows how to create conditional formatting ranges.
 * ----------------------------------------------------------------- */

using System;
using EasyXLS;
using EasyXLS.Constants;
using System.Drawing;

public class Tutorial14
{

	[STAThread]
	static void Main()
	{
		Console.WriteLine("Tutorial 14\n-----------\n");

		//Create an instance of the object that generates Excel files
		ExcelDocument xls = new ExcelDocument(1);

		//Get the table of the first sheet
		ExcelWorksheet xlsTab = ((ExcelWorksheet)xls.easy_getSheet("Sheet1"));
		ExcelTable xlsTable = xlsTab.easy_getExcelTable();

		//Insert data
		for (int i=0;i<6;i++)
		{
			for (int j=0;j<4;j++)
			{
				if((i<2)&&(j<2))
					xlsTable.easy_getCell(i, j).setValue("12");
				else
					if((j==2)&&(i<2))
					xlsTable.easy_getCell(i, j).setValue("1000");
				else
					xlsTable.easy_getCell(i, j).setValue("9");
				xlsTable.easy_getCell(i, j).setDataType(DataType.NUMERIC);
			}
		}

		//Set a conditional formatting
		xlsTab.easy_addConditionalFormatting("A1:C3", ConditionalFormatting.OPERATOR_BETWEEN, "=9", "=11", true, true, Color.Red);

		//Set a conditional formatting
		xlsTab.easy_addConditionalFormatting("A6:C6", ConditionalFormatting.OPERATOR_BETWEEN, "=COS(PI())+2", "", Color.Bisque);
		xlsTab.easy_getConditionalFormattingAt("A6:C6").getConditionAt(0).setConditionType(ConditionalFormatting.CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA);

		//Generate the file
		Console.WriteLine("Writing file C:\\Samples\\Tutorial14.xls.");
		xls.easy_WriteXLSFile("C:\\Samples\\Tutorial14.xls");

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
