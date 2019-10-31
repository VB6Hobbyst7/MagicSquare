/* ----------------------------------------------------------------
 | Tutorial 31                                                     
 |                                                                
 | This tutorial shows how to export an HTML file.	          
 -----------------------------------------------------------------*/

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{

	Console::WriteLine("Tutorial 31\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument(2);
	    
	//Set the sheet name	
	xls->easy_getSheetAt(0)->setSheetName("First tab");

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

	// Apply a predefined format to the cells.
	xlsFirstTable->easy_setRangeAutoFormat("A1:E101", gcnew ExcelAutoFormat(Styles::AUTOFORMAT_EASYXLS1));

	//Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial31.html.");
	xls->easy_WriteHTMLFile("C:\\Samples\\Tutorial31.html", "First tab");

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