/* -----------------------------------------------------------------
 * Tutorial 33
 * 
 * This tutorial shows how to set the properties of the document.
 * ----------------------------------------------------------------- */

using System;
using EasyXLS;
using EasyXLS.Constants;

public class Tutorial33
{

	[STAThread]
	static void Main()
	{
		Console.WriteLine("Tutorial 33\n-----------\n");

		//Create an instance of the object that generates Excel files
		ExcelDocument xls = new ExcelDocument(1);

		//Set the 'Subject' property
		xls.getSummaryInformation().setSubject("This is the subject");

		//Set the 'Manager' property
		xls.getDocumentSummaryInformation().setManager("This is the manager");

		//Set a custom property
		xls.getDocumentSummaryInformation().setCustomProperty("PropertyName", FileProperty.VT_NUMBER, "4");

		//Generate the file
		Console.WriteLine("Writing file C:\\Samples\\Tutorial33.xls.");
		xls.easy_WriteXLSFile("C:\\Samples\\Tutorial33.xls");

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
