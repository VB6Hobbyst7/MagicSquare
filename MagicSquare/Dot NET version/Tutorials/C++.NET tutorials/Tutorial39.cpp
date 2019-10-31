/* ----------------------------------------------------------------
| Tutorial 39
|
| This tutorial shows how to load a CSV file (we use the file
| generated in Tutorial 30), modify some data and save it to
| another file (Tutorial39.xls).
----------------------------------------------------------------- */

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;
using namespace System::Data;
using namespace System::IO;

int main()
{

	Console::WriteLine("Tutorial 39\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument();

	//Read the file
	Console::WriteLine("Reading file C:\\Samples\\Tutorial30.csv.\n");
	if (xls->easy_LoadCSVFile("C:\\Samples\\Tutorial30.csv")) 
	{
		//Set the name of the first worksheet
		xls->easy_getSheetAt(0)->setSheetName("First tab");

		//Add a gcnew worksheet and write some data
		xls->easy_addWorksheet("Second tab");
		ExcelWorksheet ^xlsSecondTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(1));
		ExcelTable ^xlsTable = xlsSecondTab->easy_getExcelTable();

        xlsTable->easy_getCell("A1")->setValue("Data added by Tutorial39");
		for (int column=0; column<5; column++)
		{
			xlsTable->easy_getCell(1, column)->setValue(String::Concat("Data ", (column + 1).ToString()));
		}


		//Generate the file
		Console::WriteLine("Writing file C:\\Samples\\Tutorial39.xls.");
		xls->easy_WriteXLSFile("C:\\Samples\\Tutorial39.xls");

		//Confirm generation
		String ^sError = xls->easy_getError();
		if (sError->Equals(""))
			Console::Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console::Write(String::Concat("\nError encountered: ", sError, "\nPress Enter to Exit..."));
	}
    else
	{
       Console::WriteLine(String::Concat("\nError reading file C:\\Samples\\Tutorial30.csv \n", xls->easy_getError(), "\nPress Enter to Exit..."));
	}
        
	//Dispose memory
    delete xls;
	
	Console::ReadLine();

	return 0;
}