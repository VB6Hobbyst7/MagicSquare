/* ----------------------------------------------------------------
 | Tutorial 03                                                     
 |                                                                
 | This tutorial shows how to create a Microsoft Excel file       
 | that has two worksheets.                                       
 -----------------------------------------------------------------*/

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;

int main()
{

	Console::WriteLine("Tutorial 03\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument(2);
	    
	//Set the sheet names	
	xls->easy_getSheetAt(0)->setSheetName("First tab");
	xls->easy_getSheetAt(1)->setSheetName("Second tab");


	//Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial03.xls.");
	xls->easy_WriteXLSFile("C:\\Samples\\Tutorial03.xls");
	
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