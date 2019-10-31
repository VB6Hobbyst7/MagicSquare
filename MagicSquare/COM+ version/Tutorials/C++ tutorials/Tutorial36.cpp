/* ----------------------------------------------------------------
 * Tutorial 36
 * 
 * This tutorial shows how to load an excel file (we use the one
 * generated in Tutorial 09), modify some data and save it to
 * another file (Tutorial36.xls).
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 36\n----------\n");

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
			printf("\nReading file: C:\\Samples\\Tutorial09.xls\n");
			if (xls->easy_LoadXLSFile("C:\\Samples\\Tutorial09.xls"))
			{
				//Get the table of the second worksheet
				EasyXLS::IExcelWorksheetPtr xlsSecondTab = (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheet("Second tab");
				EasyXLS::IExcelTablePtr xlsSecondTable = xlsSecondTab->easy_getExcelTable();

				//Write some data
				xlsSecondTable->easy_getCell_2("A1")->setValue("Data added by Tutorial36");

				char* cellValue = (char*)malloc(11*sizeof(char));
				char*  columnNumber = (char*)malloc(sizeof(char));
				for (int column=0; column<5; column++)
				{
					strcpy_s(cellValue, 6, "Data ");			
					_itoa_s(column+ 1, columnNumber, 2, 10);
					strcat_s(cellValue, 10, columnNumber);
					xlsSecondTable->easy_getCell(1, column)->setValue(cellValue);
				}
			
				
				//Generate the file
				printf("Writing file C:\\Samples\\Tutorial36.xls.");
				xls->easy_WriteXLSFile("C:\\Samples\\Tutorial36.xls");
				
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
				printf("\nError reading file C:\\Samples\\Tutorial09.xls %s\n", (LPCSTR)((_bstr_t)xls->easy_getError())); 
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
