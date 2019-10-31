/* ----------------------------------------------------------------
 * Tutorial 39
 * 
 * This tutorial shows how to load a CSV file (we use the file
 * generated in Tutorial 30), modify some data and save it to
 * another file (Tutorial39.xls).
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 39\n----------\n");

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

			//Read the file	
			printf("\nReading file: C:\\Samples\\Tutorial30.csv\n");
			if (xls->easy_LoadCSVFile("C:\\Samples\\Tutorial30.csv"))
			{
				//Set the name of the first worksheet
				xls->easy_getSheetAt(0)->setSheetName("First tab");

				//Add a new worksheet and write some data
				xls->easy_addWorksheet_2("Second tab"); 
				EasyXLS::IExcelWorksheetPtr xlsSecondTab = (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheetAt(1);
				EasyXLS::IExcelTablePtr xlsTable = xlsSecondTab->easy_getExcelTable();

				//Write some data
				xlsTable->easy_getCell_2("A1")->setValue("Data added by Tutorial39");

				char* cellValue = (char*)malloc(11*sizeof(char));
				char*  columnNumber = (char*)malloc(sizeof(char));
				for (int column=0; column<5; column++)
				{
					strcpy_s(cellValue, 6, "Data ");			
					_itoa_s(column+ 1, columnNumber, 2, 10);
					strcat_s(cellValue, 10, columnNumber);
					xlsTable->easy_getCell(1, column)->setValue(cellValue);
				}
			
				
				//Generate the file
				printf("Writing file C:\\Samples\\Tutorial39.xls.");
				xls->easy_WriteXLSFile("C:\\Samples\\Tutorial39.xls");
				
				//Confirm generation
				_bstr_t sError = xls->easy_getError();
				if (strcmp(sError, "") == 0){
					printf("\nFile successfully created. Press Enter to Exit...");
				}
				else{
					printf("\nError encountered: %s", (LPCSTR)sError); 
				}
			}
			else
			{
				printf("\nError reading file C:\\Samples\\Tutorial30.csv %s\n", (LPCSTR)((_bstr_t)xls->easy_getError())); 
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
