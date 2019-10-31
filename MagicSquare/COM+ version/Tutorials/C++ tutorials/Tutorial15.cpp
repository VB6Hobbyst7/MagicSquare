/* ----------------------------------------------------------------
 * Tutorial 15
 * 
 * This tutorial shows how to create a Hyperlink. There are 4
 * types of hyperlinks:
 *		1 - to an URL;
 * 		2 - to a FILE;
 * 		3 - to a UNC;
 * 		4 - to a CELL in the same file;
 * 
 * The link can be placed over multiple cells.
 * 
 * Every type of hyperlink accepts a tool tip description.
 * 
 * Every type of hyperlink accepts a text mark. A text mark is a
 * link inside the file. Examples:
 * 		http://www.mysite.com/index.html#Chapter3
 * 		c:\myfile.xls#Sheet2!D3
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 15\n----------\n");

	HRESULT hr;

	//Initialize COM
	hr = CoInitialize(0);

	// Use the SUCCEEDED macro and get a pointer to the interface
	if(SUCCEEDED(hr))
	{
		//Create a pointer to the interface that generates Excel files
		EasyXLS::IExcelDocumentPtr xls;
		hr = CoCreateInstance(__uuidof(EasyXLS::ExcelDocument),
		NULL,
		CLSCTX_ALL,
		__uuidof(EasyXLS::IExcelDocument),
		(void**) &xls) ;

		if(SUCCEEDED(hr)){

			//Create the worksheets 
			xls->easy_addWorksheet_2("First tab");
			xls->easy_addWorksheet_2("Second tab");

			//Get the table of the first sheet
			EasyXLS::IExcelWorksheetPtr xlsTab1 = (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheetAt(0);	
			EasyXLS::IExcelWorksheetPtr xlsTab2 = (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheetAt(1);	
			
			//Create the hyperlink to an URL
			xlsTab1->easy_addHyperlink_3(HYPERLINKTYPE_URL, "http://www.euoutsourcing.com", "Link to URL", "B2:E2");

			//Create the hyperlink to a FILE
			xlsTab1->easy_addHyperlink_3(HYPERLINKTYPE_FILE, "c:\\myfile.xls", "Link to file", "B3");

			//Create the hyperlink to an UNC
			xlsTab1->easy_addHyperlink_3(HYPERLINKTYPE_UNC, "\\\\computerName\\Folder\\file.txt", "Link to UNC", "B4:D4");

			//Create the hyperlink to a CELL
			xlsTab1->easy_addHyperlink_3(HYPERLINKTYPE_CELL, "'Second tab'!D3", "Link to CELL", "B5");

			//Creating a name for the second sheet
			xlsTab2->easy_addName_2("Name", "=Second tab!$A$1:$A$4");
			
			//Create the hyperlink to a name
			xlsTab1->easy_addHyperlink_3(HYPERLINKTYPE_CELL, "Name", "Link to a name", "B6");


			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial15.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial15.xls");
			
			//Confirm generation
			_bstr_t sError = xls->easy_getError();
			if (strcmp(sError, "") == 0){
				printf("\nFile successfully created. Press Enter to Exit...");
			}
			else{
				printf("\nError encountered: %s", (LPCSTR)sError); 
			}
			
			//Dispose memory
			xls->Dispose();
		}
		else{
			printf("Object is not available!");
		}
	}
	else{
		printf("COM can't be initialized!");
	}

	// Uninitialize COM
	CoUninitialize();

	_getch();
	return 0;
}