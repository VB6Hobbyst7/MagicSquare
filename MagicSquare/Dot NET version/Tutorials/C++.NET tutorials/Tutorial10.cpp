/* -----------------------------------------------------------------
 | Tutorial 10                                                     |
 |                                                                 |
 | This tutorial shows how to merge a cell range.                  |
  -----------------------------------------------------------------*/

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;

int main()
{

	Console::WriteLine("Tutorial 10\n----------\n");

	// Create an instance of the object that generates Excel files
	ExcelDocument ^xls = gcnew ExcelDocument(1);

	// Get the table of the first sheet
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheet("Sheet1"));
	ExcelTable ^xlsTable = xlsFirstTab->easy_getExcelTable();

	// Merging cells
	xlsTable->easy_mergeCells("A1:C3");    


	// Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial10.xls.");
	xls->easy_WriteXLSFile("C:\\Samples\\Tutorial10.xls");

	// Confirm generation
	String ^sError = xls->easy_getError();
	if (sError->Equals(""))
		Console::Write("\nFile successfully created. Press Enter to Exit...");
	else
		Console::Write(String::Concat("\nError encountered: ", sError, "\nPress Enter to Exit..."));

	// Dispose memory
	delete xls;

	Console::ReadLine();

	return 0;
}