/* ----------------------------------------------------------------
| Tutorial 41
|
| This tutorial shows how to load an XML file (we use the file
| generated in Tutorial 32), modify some data and save it to
| another file (Tutorial41.xls).
----------------------------------------------------------------- */

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;
using namespace System::Data;
using namespace System::IO;

int main()
{

	Console::WriteLine("Tutorial 41\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument();

	//Read the file
	Console::WriteLine("Reading file C:\\Samples\\Tutorial32.xml.\n");
	if (xls->easy_LoadXMLSpreadsheetFile("C:\\Samples\\Tutorial32.xml")) 
	{
		//Get the table of the second worksheet
		ExcelWorksheet ^xlsSecondTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(1));
		ExcelTable ^xlsTable = xlsSecondTab->easy_getExcelTable();

        xlsTable->easy_getCell("A1")->setValue("Data added by Tutorial41");
		for (int column=0; column<5; column++)
		{
			xlsTable->easy_getCell(1, column)->setValue(String::Concat("Data ", (column + 1).ToString()));
		}


		//Generate the file
		Console::WriteLine("Writing file C:\\Samples\\Tutorial41.xls.");
		xls->easy_WriteXLSFile("C:\\Samples\\Tutorial41.xls");

		//Confirm generation
		String ^sError = xls->easy_getError();
		if (sError->Equals(""))
			Console::Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console::Write(String::Concat("\nError encountered: ", sError, "\nPress Enter to Exit..."));
	}
    else
	{
       Console::WriteLine(String::Concat("\nError reading file C:\\Samples\\Tutorial32.xml \n", xls->easy_getError(), "\nPress Enter to Exit..."));
	}
        
	//Dispose memory
    delete xls;
	
	Console::ReadLine();

	return 0;
}