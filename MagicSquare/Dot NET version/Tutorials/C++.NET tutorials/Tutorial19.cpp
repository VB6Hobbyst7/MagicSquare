/* ----------------------------------------------------------------
| Tutorial 19													
|																	
| This tutorial shows how to create a Microsoft Excel file
| that has two worksheets. The first one is full with data
| and the first cell of the second row contains Rich Text Format.
-----------------------------------------------------------------*/

#include "stdafx.h"
#include <conio.h>
#include <stdio.h>

using namespace System;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{


	Console::WriteLine("Tutorial 19\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument(2);
	
	//Set the sheet names	
	xls->easy_getSheetAt(0)->setSheetName("First tab");
	xls->easy_getSheetAt(1)->setSheetName("Second tab");
	
	//Get the table of the first worksheet
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(0));
	ExcelTable ^xlsFirstTable = xlsFirstTab->easy_getExcelTable();

		//Create the string used to set the RTF in cell
		String ^ sFormattedValue = "This is <b>bold</b>.";
		sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <i>italic</i>.");
		sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <u>underline</u>.");
		sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <underline double>double underline</underline double>.");
		sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <font color=red>red</font>.");
		sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <font color=rgb(255,0,0)>red</font> too.");
		sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <font face=\"Arial Black\">Arial Black</font>.");
		sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <font size=15pt>size 15</font>.");
		sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <s>strikethrough</s>.");
		sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <sup>superscript</sup>.");
		sFormattedValue = String::Concat(sFormattedValue ,"\nThis is <sub>subscript</sub>.");
		sFormattedValue = String::Concat(sFormattedValue ,"\n<b>This</b> <i>is</i> <font color=red face=\"Arial Black\" size=15pt><underline double>formatted</underline double></font> <s>text</s>.");
		

		//Set the formatted value
		xlsFirstTable->easy_getCell(1, 0)->setHTMLValue(sFormattedValue); 
		xlsFirstTable->easy_getCell(1, 0)->setDataType(DataType::STRING);
		xlsFirstTable->easy_getCell(1, 0)->setWrap(true); 
		xlsFirstTable->easy_getRowAt(1)->setHeight(250);
		xlsFirstTable->easy_getColumnAt(0)->setWidth(250);
		



	//Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial19.xls.");
	xls->easy_WriteXLSFile("C:\\Samples\\Tutorial19.xls");

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