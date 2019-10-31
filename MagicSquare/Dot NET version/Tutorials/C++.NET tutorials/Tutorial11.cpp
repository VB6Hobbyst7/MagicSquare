/* ----------------------------------------------------------------
 | Tutorial 11                                                    |
 |                                                                |
 | This tutorial shows how to create a Microsoft Excel file       |
 | that has a formula.                                            |
 -----------------------------------------------------------------*/

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;

int main()
{

	Console::WriteLine("Tutorial 11\n----------\n");

	//Create an instance of the object that generates Excel files
	ExcelDocument ^xls = gcnew ExcelDocument();

	//Add one worksheet
	xls->easy_addWorksheet("Formula");

	//Get the table, populate the sheet and set a formula
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheet("Formula"));
	ExcelTable ^xlsTable = xlsFirstTab->easy_getExcelTable();
	xlsTable->easy_getCell("A1")->setValue("1");
	xlsTable->easy_getCell("A2")->setValue("2");
	xlsTable->easy_getCell("A3")->setValue("3");
	xlsTable->easy_getCell("A4")->setValue("4");
	xlsTable->easy_getCell("A6")->setValue("=SUM(A1:A4)");


	//Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial11.xls.");
	xls->easy_WriteXLSFile("C:\\Samples\\Tutorial11.xls");

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