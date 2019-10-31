/* ----------------------------------------------------------------------------------
 | Tutorial 29                                                     
 |                                                                
 | This tutorial shows how to export a XLSB file that has multiple sheets in C++.NET. 
 | The first sheet is filled with data.	          
 --------------------------------------------------------------------------------- */

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{

	Console::WriteLine("Tutorial 29\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument(2);
	    
	//Set the sheet name	
	xls->easy_getSheetAt(0)->setSheetName("First tab");
	xls->easy_getSheetAt(1)->setSheetName("Second tab");

	//Get the table of the first worksheet
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(0));
	ExcelTable ^xlsFirstTable = xlsFirstTab->easy_getExcelTable();

	//Add the cells for header
	for (int column=0; column<5; column++)
	{
		xlsFirstTable->easy_getCell(0,column)->setValue(String::Concat("Column ",(column + 1).ToString())); 
		xlsFirstTable->easy_getCell(0,column)->setDataType(DataType::STRING);
	}

	//Add the cells for data
	for (int row=0; row<100; row++)
	{
		for (int column=0; column<5; column++)
		{
			xlsFirstTable->easy_getCell(row+1,column)->setValue(String::Concat("Data ", (row + 1).ToString(), ", ", (column + 1).ToString())); 
			xlsFirstTable->easy_getCell(row+1,column)->setDataType(DataType::STRING);
		}
	}
	//Set column widths
	xlsFirstTable->setColumnWidth(0, 70);
	xlsFirstTable->setColumnWidth(1, 100);
	xlsFirstTable->setColumnWidth(2, 70);
	xlsFirstTable->setColumnWidth(3, 100);
	xlsFirstTable->setColumnWidth(4, 70);


	//Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial29.xlsb.");
	xls->easy_WriteXLSBFile("C:\\Samples\\Tutorial29.xlsb");

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