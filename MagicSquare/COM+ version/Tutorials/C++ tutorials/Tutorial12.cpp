/* ----------------------------------------------------------------
 * Tutorial 12
 * 
 * This tutorial shows how to create a Microsoft Excel file
 * that has two worksheets. The second one contains a named
 * range.
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 12\n----------\n");

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

			//Get the table of the second worksheet and populate the sheet
			EasyXLS::IExcelWorksheetPtr xlsSecondTab = (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheetAt(1);
			EasyXLS::IExcelTablePtr xlsSecondTable = xlsSecondTab->easy_getExcelTable();
			xlsSecondTable->easy_getCell_2("A1")->setValue("Range data 1");
			xlsSecondTable->easy_getCell_2("A2")->setValue("Range data 2");
			xlsSecondTable->easy_getCell_2("A3")->setValue("Range data 3");
			xlsSecondTable->easy_getCell_2("A4")->setValue("Range data 4");

			//Create a named range
			xlsSecondTab->easy_addName_2("Range", "='Second tab'!$A$1:$A$4");


			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial12.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial12.xls");
			
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