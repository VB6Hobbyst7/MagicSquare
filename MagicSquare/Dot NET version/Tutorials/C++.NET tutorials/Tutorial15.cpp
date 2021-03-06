/* ----------------------------------------------------------------
 | Tutorial 15                                                    
 |                                                                
 | This tutorial shows how to create a Hyperlink. There are 4     
 | types of hyperlinks:											
 |		1 - to an URL;											
 |		2 - to a FILE;											 
 |		3 - to a UNC;											 
 |		4 - to a CELL in the same file;							  
 |																  
 | The link can be placed over multiple cells.					  
 |																  
 | Every type of hyperlink accepts a tool tip description.		  
 |																  
 | Every type of hyperlink accepts a text mark. A text mark is a  
 | link inside the file. Examples:								  
 |		http://www.mysite.com/index.html#Chapter3				  
 |		c:\myfile.xls#Sheet2!D3									 
 -----------------------------------------------------------------*/

#include "stdafx.h"

using namespace System;
using namespace EasyXLS;

int main()
{

	Console::WriteLine("Tutorial 15\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument(2);
	
	//Set the sheet names
	ExcelWorksheet ^xlsTab1 = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(0));
	ExcelWorksheet ^xlsTab2 = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(1));
	xlsTab1->setSheetName("First tab");
	xlsTab2->setSheetName("Second tab");

	//Create the hyperlink to an URL
	xlsTab1->easy_addHyperlink(EasyXLS::Constants::HyperlinkType::URL, "http://www.euoutsourcing.com", "Link to URL", "B2:E2");

	//Create the hyperlink to a FILE
	xlsTab1->easy_addHyperlink(EasyXLS::Constants::HyperlinkType::FILE, "c:\\myfile.xls", "Link to file", "B3");

	//Create the hyperlink to an UNC
	xlsTab1->easy_addHyperlink(EasyXLS::Constants::HyperlinkType::UNC, "\\\\computerName\\Folder\\file.txt", "Link to UNC", "B4:D4");

	//Create the hyperlink to a CELL
	xlsTab1->easy_addHyperlink(EasyXLS::Constants::HyperlinkType::CELL, "'Second tab'!D3", "Link to CELL", "B5");

	//Creating a name for the second sheet
	xlsTab2->easy_addName("Name", "=Second tab!$A$1:$A$4");
	
	//Create the hyperlink to a name
	xlsTab1->easy_addHyperlink(EasyXLS::Constants::HyperlinkType::CELL, "Name", "Link to a name", "B6");


	//Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial15.xls.");
	xls->easy_WriteXLSFile("C:\\Samples\\Tutorial15.xls");

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