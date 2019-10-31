/* ----------------------------------------------------------------
| Tutorial 37
|
| This tutorial shows how to load a XLSX file (we use the file
| generated in Tutorial 28), modify some data and save it to
| another file (Tutorial37.xlsx).
----------------------------------------------------------------- */

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;
using namespace System::Data;
using namespace System::IO;

int main()
{

	Console::WriteLine("Tutorial 37\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument();

	//Read the file
	Console::WriteLine("Reading file C:\\Samples\\Tutorial28.xlsx.\n");
	if (xls->easy_LoadXLSXFile("C:\\Samples\\Tutorial28.xlsx")) 
	{
		//Get the table of the second worksheet
		ExcelWorksheet ^xlsSecondTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheet("Second tab"));
		ExcelTable ^xlsSecondTable = xlsSecondTab->easy_getExcelTable();

        xlsSecondTable->easy_getCell("A1")->setValue("Data added by Tutorial37");
		for (int column=0; column<5; column++)
		{
			xlsSecondTable->easy_getCell(1, column)->setValue(String::Concat("Data ", (column + 1).ToString()));
		}


		//Generate the file
		Console::WriteLine("Writing file C:\\Samples\\Tutorial37.xlsx.");
		xls->easy_WriteXLSXFile("C:\\Samples\\Tutorial37.xlsx");

		//Confirm generation
		String ^sError = xls->easy_getError();
		if (sError->Equals(""))
			Console::Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console::Write(String::Concat("\nError encountered: ", sError, "\nPress Enter to Exit..."));
	}
    else
	{
       Console::WriteLine(String::Concat("\nError reading file C:\\Samples\\Tutorial28.xlsx \n", xls->easy_getError(), "\nPress Enter to Exit..."));
	}
        
	//Dispose memory
    delete xls;
	
	Console::ReadLine();

	return 0;
}