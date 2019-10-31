/* ----------------------------------------------------------------
 * Tutorial 38
 * 
 * This tutorial shows how to load an XLSB file (we use the file
 * generated in Tutorial 29), modify some data and save it to
 * another file (Tutorial38.xlsb).
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 38\n----------\n");

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
			printf("\nReading file: C:\\Samples\\Tutorial29.xlsb\n");
			if (xls->easy_LoadXLSBFile("C:\\Samples\\Tutorial29.xlsb"))
			{
				//Get the table of the second worksheet
				EasyXLS::IExcelWorksheetPtr xlsSecondTab = (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheetAt(1);
				//Write some data
				EasyXLS::IExcelTablePtr xlsTable = xlsSecondTab->easy_getExcelTable();

				//Write some data
				xlsTable->easy_getCell_2("A1")->setValue("Data added by Tutorial38");

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
				printf("Writing file C:\\Samples\\Tutorial38.xlsb.");
				xls->easy_WriteXLSBFile("C:\\Samples\\Tutorial38.xlsb");
				
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
				printf("\nError reading file C:\\Samples\\Tutorial29.xlsb %s\n", (LPCSTR)((_bstr_t)xls->easy_getError())); 
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
