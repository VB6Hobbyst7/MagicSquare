/* ----------------------------------------------------------------
 * Tutorial 11
 * 
 * This tutorial shows how to create a Microsoft Excel file
 * that has a formula.
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 11\n----------\n");

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

			//Add one worksheet
			xls->easy_addWorksheet_2("Formula");

			//Get the table, populate the sheet and set a formula
			EasyXLS::IExcelWorksheetPtr xlsFirstTab= (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheet("Formula");
			EasyXLS::IExcelTablePtr xlsTable = xlsFirstTab->easy_getExcelTable();
			xlsTable->easy_getCell_2("A1")->setValue("1");
			xlsTable->easy_getCell_2("A2")->setValue("2");
			xlsTable->easy_getCell_2("A3")->setValue("3");
			xlsTable->easy_getCell_2("A4")->setValue("4");
			xlsTable->easy_getCell_2("A6")->setValue("=SUM(A1:A4)");


			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial11.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial11.xls");
			
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