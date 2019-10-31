/* ----------------------------------------------------------------
 | Tutorial 09                                                    
 |                                                                
 | This tutorial shows how to create a Microsoft Excel file       
 | that has two worksheets. The first one is full with data       
 | and the cells are formatted. The column header has comments.   
 | The first worksheet has header & footer. The print options are 
 | set for the first worksheet.                                   
 -----------------------------------------------------------------*/

#include "stdafx.h"

using namespace System;
using namespace System::Drawing;
using namespace EasyXLS;
using namespace EasyXLS::Constants;

int main()
{

	Console::WriteLine("Tutorial 09\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument(2);
	
	//Set the sheet names	
	xls->easy_getSheetAt(0)->setSheetName("First tab");
	xls->easy_getSheetAt(1)->setSheetName("Second tab");

	//Lock the first tab
	xls->easy_getSheetAt(0)->setSheetProtected(true);

	//Get the table of the first worksheet
	ExcelWorksheet ^xlsFirstTab = safe_cast<ExcelWorksheet^>(xls->easy_getSheetAt(0));
	ExcelTable ^xlsFirstTable = xlsFirstTab->easy_getExcelTable();

	//Create the style for the header
	ExcelStyle ^xlsStyleHeader = gcnew ExcelStyle("Verdana", 8, true, true, Color::Yellow);		
	xlsStyleHeader->setBackground(Color::Black);
	xlsStyleHeader->setBorderColors(Color::Gray, Color::Gray, Color::Gray, Color::Gray);
	xlsStyleHeader->setBorderStyles(Border::BORDER_MEDIUM, Border::BORDER_MEDIUM, Border::BORDER_MEDIUM, Border::BORDER_MEDIUM);	
	xlsStyleHeader->setHorizontalAlignment(Alignment::ALIGNMENT_CENTER);
	xlsStyleHeader->setVerticalAlignment(Alignment::ALIGNMENT_BOTTOM);
	xlsStyleHeader->setWrap(true);
	xlsStyleHeader->setDataType(DataType::STRING);

	//Add the cells for header
	for (int column=0; column<5; column++)
	{
		xlsFirstTable->easy_getCell(0,column)->setValue(String::Concat("Column ",(column + 1).ToString())); 
		xlsFirstTable->easy_getCell(0,column)->setStyle(xlsStyleHeader);

		//Add comment
		xlsFirstTable->easy_getCell(0, column)->setComment(String::Concat("This is column no ",(column + 1).ToString()));
	}
	xlsFirstTable->easy_getRowAt(0)->setHeight(30);

	//Add the cells for data
	for (int row=0; row<100; row++)
	{
		for (int column=0; column<5; column++)
		{
			xlsFirstTable->easy_getCell(row+1,column)->setValue(String::Concat("Data ", (row + 1).ToString(), ", ", (column + 1).ToString())); 
		}
	}

	//Create a style for cells
	ExcelStyle ^xlsStyleData = gcnew ExcelStyle();
	xlsStyleData->setHorizontalAlignment(Alignment::ALIGNMENT_LEFT);
	xlsStyleData->setForeground(Color::DarkGray);
	xlsStyleData->setWrap(false);
	xlsStyleData->setDataType(DataType::STRING);
	xlsStyleData->setLocked(true);
	xlsFirstTable->easy_setRangeStyle("A2:E101", xlsStyleData);

	//Set column widths
	xlsFirstTable->setColumnWidth(0, 70);
	xlsFirstTable->setColumnWidth(1, 100);
	xlsFirstTable->setColumnWidth(2, 70);
	xlsFirstTable->setColumnWidth(3, 100);
	xlsFirstTable->setColumnWidth(4, 70);

	//Add headers for the first worksheet
	xlsFirstTab->easy_getHeaderAt(Header::POSITION_CENTER)->InsertSingleUnderline();
	xlsFirstTab->easy_getHeaderAt(Header::POSITION_CENTER)->InsertFile();
	xlsFirstTab->easy_getHeaderAt(Header::POSITION_CENTER)->InsertValue(" - How to create header and footer");

	xlsFirstTab->easy_getHeaderAt(Header::POSITION_RIGHT)->InsertDate();
	xlsFirstTab->easy_getHeaderAt(Header::POSITION_RIGHT)->InsertValue(" ");
	xlsFirstTab->easy_getHeaderAt(Header::POSITION_RIGHT)->InsertTime();

	//Add footer for the first worksheet
	xlsFirstTab->easy_getFooterAt(Footer::POSITION_CENTER)->InsertPage();
	xlsFirstTab->easy_getFooterAt(Footer::POSITION_CENTER)->InsertValue(" of ");
	xlsFirstTab->easy_getFooterAt(Footer::POSITION_CENTER)->InsertPages();

	//Set Page Setup options
	ExcelPageSetup ^xlsPageSetup = xlsFirstTab->easy_getPageSetup();
	xlsPageSetup->easy_setPrintArea("A1:E101");
	xlsPageSetup->easy_setRowsToRepeatAtTop("$1:$1");
	xlsPageSetup->setCenterHorizontally(true);
	xlsPageSetup->setOrientation(PageSetup::ORIENTATION_PORTRAIT);
	xlsPageSetup->setPageOrder(PageSetup::PAGE_ORDER_DOWN_THEN_OVER);
	xlsPageSetup->setPaperSize(PageSetup::PAPER_SIZE_A4);
	xlsPageSetup->setPrintComments(PageSetup::COMMENTS_AT_END_OF_SHEET);
	xlsPageSetup->setPrintGridlines(true);
	xlsFirstTable->easy_insertPageBreakAtRow(21);
	xlsFirstTable->easy_insertPageBreakAtRow(41);
	xlsFirstTable->easy_insertPageBreakAtRow(61);
	xlsFirstTable->easy_insertPageBreakAtRow(81);
	xlsFirstTab->setPageBreakPreview(true);


	//Generate the file
	Console::WriteLine("Writing file C:\\Samples\\Tutorial09.xls.");
	xls->easy_WriteXLSFile("C:\\Samples\\Tutorial09.xls");

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