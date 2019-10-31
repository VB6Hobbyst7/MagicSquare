/* ----------------------------------------------------------------
 * Tutorial 16
 * 
 * This tutorial shows how to create a Microsoft Excel file
 * that has two worksheets. The first one has an image.
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 16\n----------\n");

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

			//Create the image
			EasyXLS::IExcelWorksheetPtr xlsFirstTab = (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheetAt(0);
			xlsFirstTab->easy_addImage_5("C:\\Samples\\EasyXLSLogo.JPG", "A1");
			

			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial16.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial16.xls");
			
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