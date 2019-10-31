/* ----------------------------------------------------------------
 | Tutorial 12                                                    |
 |                                                                |
 | This tutorial shows how to create a Microsoft Excel file       |
 | that has two worksheets. The second one contains a named       |
 | range.                                                         |
 -----------------------------------------------------------------*/

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;

int main()
{

	Console::WriteLine("Tutorial 12\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument(2);
	
	//Set the sheet names	
	xls->easy_getSheetAt(0)->setSheetName("First tab");
	xls->easy_getSheetAt(1)->setSheetName("Second tab");

	//Get the table of the second worksheet and populate the sheet
	ExcelWorksheet ^xlsSecondTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(1));
	ExcelTable ^xlsSecondTable = xlsSecondTab->easy_getExcelTable();
	xlsSecondTable->easy_getCell("A1")->setValue("Range data 1");
	xlsSecondTable->easy_getCell("A2")->setValue("Range data 2");
	xlsSecondTable->easy_getCell("A3")->setValue("Range data 3");
	xlsSecondTable->easy_getCell("A4")->setValue("Range data 4");

	//Create a named range
	xlsSecondTab->easy_addName("Range", "='Second tab'!$A$1:$A$4");


	//Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial12.xls.");
	xls->easy_WriteXLSFile("C:\\Samples\\Tutorial12.xls");

	//Confirm generation
	String ^sError = xls->easy_getError();
	if (sError->Equals(""))
		Console::Write("\nFile successfully created. Press Enter to Exit...");
	else
		Console::Write(String::Concat("\nError encountered: ", sError, "\nPress Enter to Exit..."));

	//Dispose memory
	delete xls;

	Console::ReadLine();

	return 0;
}