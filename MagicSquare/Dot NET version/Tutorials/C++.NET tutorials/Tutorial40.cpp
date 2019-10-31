/* ----------------------------------------------------------------
| Tutorial 40
|
| This tutorial shows how to load an HTML file (we use the file
| generated in Tutorial 31), modify some data and save it to
| another file (Tutorial40.xls).
----------------------------------------------------------------- */

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;
using namespace System::Data;
using namespace System::IO;

int main()
{

	Console::WriteLine("Tutorial 40\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument();

	//Read the file
	Console::WriteLine("Reading file C:\\Samples\\Tutorial31.html.\n");
	if (xls->easy_LoadHTMLFile("C:\\Samples\\Tutorial31.html")) 
	{
		//Set the name of the first worksheet
		xls->easy_getSheetAt(0)->setSheetName("First tab");

		//Add a gcnew worksheet and write some data
		xls->easy_addWorksheet("Second tab");
		ExcelWorksheet ^xlsSecondTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(1));
		ExcelTable ^xlsTable = xlsSecondTab->easy_getExcelTable();

        xlsTable->easy_getCell("A1")->setValue("Data added by Tutorial40");
		for (int column=0; column<5; column++)
		{
			xlsTable->easy_getCell(1, column)->setValue(String::Concat("Data ", (column + 1).ToString()));
		}


		//Generate the file
		Console::WriteLine("Writing file C:\\Samples\\Tutorial40.xls.");
		xls->easy_WriteXLSFile("C:\\Samples\\Tutorial40.xls");

		//Confirm generation
		String ^sError = xls->easy_getError();
		if (sError->Equals(""))
			Console::Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console::Write(String::Concat("\nError encountered: ", sError, "\nPress Enter to Exit..."));
	}
    else
	{
       Console::WriteLine(String::Concat("\nError reading file C:\\Samples\\Tutorial31.html \n", xls->easy_getError(), "\nPress Enter to Exit..."));
	}
        
	//Dispose memory
    delete xls;
	
	Console::ReadLine();

	return 0;
}