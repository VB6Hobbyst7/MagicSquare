/* -----------------------------------------------------------------
 * Tutorial 20
 * 
 * This tutorial shows how to create a Microsoft Excel file 
 * that has AutoFilter.
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 20\n----------\n");

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
			xls->easy_addWorksheet_2("Sheet1");

			//Get the table of the first worksheet
			EasyXLS::IExcelWorksheetPtr xlsTab= (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheet("Sheet1");
			EasyXLS::IExcelTablePtr xlsTable = xlsTab->easy_getExcelTable();
			
			//Add the cells for header
			char* cellValue = (char*)malloc(11*sizeof(char));
			char*  columnNumber = (char*)malloc(sizeof(char));
			for (int column=0; column<5; column++)
			{
				strcpy_s(cellValue, 8, "Column ");			
				_itoa_s(column+ 1, columnNumber, 2, 10);
				strcat_s(cellValue, 10, columnNumber);
				xlsTable->easy_getCell(0,column)->setValue(cellValue); 
				xlsTable->easy_getCell(0,column)->setDataType(DATATYPE_STRING);
			}

			//Add the cells for data
			char*  rowNumber = (char*)malloc(sizeof(char));
			for (int row=0; row<100; row++)
			{
				for (int column=0; column<5; column++)
				{
					strcpy_s(cellValue, 6, "Data ");	
					_itoa_s(column+ 1, columnNumber, 2, 10);
					_itoa_s(row + 1, rowNumber, 4, 10);

					strcat_s(cellValue, 10, rowNumber);
					strcat_s(cellValue, 12, ", ");
					strcat_s(cellValue, 13, columnNumber);

					xlsTable->easy_getCell(row+1,column)->setValue(cellValue); 
					xlsTable->easy_getCell(row+1,column)->setDataType(DATATYPE_STRING);
				}
			}

			//Add AutoFilter
			EasyXLS::IExcelFilterPtr xlsFilter = (EasyXLS::IExcelFilterPtr)xlsTab->easy_getFilter();
			xlsFilter->setAutoFilter_2("A1:E1");
			
			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial20.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial20.xls");
			
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