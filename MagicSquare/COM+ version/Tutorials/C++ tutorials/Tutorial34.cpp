/* ----------------------------------------------------------------
 * Tutorial 34
 * 
 * This tutorial shows how to read values from the active sheet
 * of an excel file (the file generated in Tutorial 09).
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 34\n----------\n");

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
			EasyXLS::IListPtr  rows = xls->easy_ReadXLSActiveSheet_AsList("C:\\Samples\\Tutorial09.xls");
		
			//Confirm reading	
			_bstr_t sError = xls->easy_getError();
			if (strcmp(sError, "") == 0){
			
				//Display the values
				for ( int rowIndex=0; rowIndex<rows->size(); rowIndex++)
				{
					EasyXLS::IListPtr 	row = (EasyXLS::IListPtr) rows->elementAt(rowIndex);
					for (int cellIndex=0; cellIndex<row->size(); cellIndex++)
					{
						printf("At row %d, column %d the value is '%s'\n", (rowIndex+ 1), (cellIndex+ 1), (LPCSTR)((_bstr_t)row->elementAt(cellIndex)));
					}
				}
                printf("\nPress Enter to exit ...");
			}
			else
			{
				printf("\nError reading file C:\\Samples\\Tutorial09.xls %s\n", (LPCSTR)sError); 
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
