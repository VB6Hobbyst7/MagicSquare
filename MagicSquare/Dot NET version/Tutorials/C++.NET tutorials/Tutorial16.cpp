/* ----------------------------------------------------------------
| Tutorial 16                                                    |
|                                                                |
| This tutorial shows how to create a Microsoft Excel file       |
| that has two worksheets. The first one has an image.           |
-----------------------------------------------------------------*/

#include "stdafx.h"
#include <conio.h>
#include <stdio.h>

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{


	Console::WriteLine("Tutorial 16\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument(2);
	
	//Set the sheet names	
	xls->easy_getSheetAt(0)->setSheetName("First tab");
	xls->easy_getSheetAt(1)->setSheetName("Second tab");
	
	//Create the image
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(0));
	xlsFirstTab->easy_addImage("C:\\Samples\\EasyXLSLogo.JPG", "A1");

	//Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial16.xls.");
	xls->easy_WriteXLSFile("C:\\Samples\\Tutorial16.xls");
	
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