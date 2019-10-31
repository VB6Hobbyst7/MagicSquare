/* ----------------------------------------------------------------
| Tutorial 20													
|																	
| This tutorial shows how to create a Microsoft Excel file 
| that has AutoFilter.

-----------------------------------------------------------------*/

#include "stdafx.h"
#include <conio.h>
#include <stdio.h>

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{
	Console::WriteLine("Tutorial 20\n----------\n");

	//Create an instance of the object that generates Excel files, having 1 sheet
	ExcelDocument ^xls = gcnew ExcelDocument(1);
	
	//Get the table of the first worksheet
	ExcelWorksheet ^xlsTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheet("Sheet1"));
	ExcelTable ^xlsTable = xlsTab->easy_getExcelTable();

	//Add the cells for header
	for (int column=0; column<5; column++)
	{
		xlsTable->easy_getCell(0,column)->setValue(String::Concat("Column ",(column + 1).ToString())); 
		xlsTable->easy_getCell(0,column)->setDataType(DataType::STRING);
	}

	//Add the cells for data
	for (int row=0; row<100; row++)
	{
		for (int column=0; column<5; column++)
		{
			xlsTable->easy_getCell(row+1,column)->setValue(String::Concat("Data ", (row + 1).ToString(), ", ", (column + 1).ToString())); 
			xlsTable->easy_getCell(row+1,column)->setDataType(DataType::STRING);
		}
	}

	//Add AutoFilter
	ExcelFilter ^xlsFilter = xlsTab->easy_getFilter();
	xlsFilter->setAutoFilter("A1:E1");

	//Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial20.xls.");
	xls->easy_WriteXLSFile("C:\\Samples\\Tutorial20.xls");

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