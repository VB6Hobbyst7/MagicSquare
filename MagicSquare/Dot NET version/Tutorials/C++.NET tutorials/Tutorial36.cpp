/* ----------------------------------------------------------------
| Tutorial 36
|
| This tutorial shows how to load an excel file (we use the one
| generated in Tutorial 09), modify some data and save it to
| another file (Tutorial36.xls).
----------------------------------------------------------------- */

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;
using namespace System::Data;
using namespace System::IO;

int main()
{

	Console::WriteLine("Tutorial 36\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument();

	//Read the file
	Console::WriteLine("Reading file C:\\Samples\\Tutorial09.xls.\n");
	if (xls->easy_LoadXLSFile("C:\\Samples\\Tutorial09.xls")) 
	{
		//Get the table of the second worksheet
		ExcelWorksheet ^xlsSecondTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(1));
		ExcelTable ^xlsSecondTable = xlsSecondTab->easy_getExcelTable();

		//Write some data
        xlsSecondTable->easy_getCell("A1")->setValue("Data added by Tutorial36");
		for (int column=0; column<5; column++)
		{
			xlsSecondTable->easy_getCell(1, column)->setValue(String::Concat("Data ", (column + 1).ToString()));
		}


		//Generate the file
		Console::WriteLine("Writing file C:\\Samples\\Tutorial36.xls.");
		xls->easy_WriteXLSFile("C:\\Samples\\Tutorial36.xls");

		//Confirm generation
		String ^sError = xls->easy_getError();
		if (sError->Equals(""))
			Console::Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console::Write(String::Concat("\nError encountered: ", sError, "\nPress Enter to Exit..."));
	}
    else
	{
       Console::WriteLine(String::Concat("\nError reading file C:\\Samples\\Tutorial09.xls \n", xls->easy_getError(), "\nPress Enter to Exit..."));
	}
        
	//Dispose memory
    delete xls;
	
	Console::ReadLine();

	return 0;
}