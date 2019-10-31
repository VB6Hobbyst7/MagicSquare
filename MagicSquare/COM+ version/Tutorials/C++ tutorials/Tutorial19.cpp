/* -----------------------------------------------------------------
 * Tutorial 19
 * 
 * This tutorial shows how to create a Microsoft Excel file
 * that has two worksheets. The first one is full with data 
 * and the first cell of the second row contains Rich Text Format.
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 19\n----------\n");

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

			//Get the table of the first worksheet
			EasyXLS::IExcelWorksheetPtr xlsFirstTab= (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheetAt(0);
			EasyXLS::IExcelTablePtr xlsFirstTable = xlsFirstTab->easy_getExcelTable();
			

			//Create the string used to set the RTF in cell
			char* sFormattedValue = (char*)malloc(536*sizeof(char));
			strcpy_s(sFormattedValue, 21, "This is <b>bold</b>.");		
			strcpy_s(sFormattedValue , 24, "\nThis is <i>italic</i>.");
			strcpy_s(sFormattedValue , 27, "\nThis is <u>underline</u>.");
			strcpy_s(sFormattedValue , 64, "\nThis is <underline double>double underline</underline double>.");
			strcpy_s(sFormattedValue , 37, "\nThis is <font color=red>red</font>.");
			strcpy_s(sFormattedValue , 50, "\nThis is <font color=rgb(255,0,0)>red</font> too.");
			strcpy_s(sFormattedValue , 56, "\nThis is <font face=\"Arial Black\">Arial Black</font>.");
			strcpy_s(sFormattedValue , 41, "\nThis is <font size=15pt>size 15</font>.");
			strcpy_s(sFormattedValue , 31, "\nThis is <s>strikethrough</s>.");
			strcpy_s(sFormattedValue , 33, "\nThis is <sup>superscript</sup>.");
			strcpy_s(sFormattedValue , 31, "\nThis is <sub>subscript</sub>.");
			strcpy_s(sFormattedValue , 138, "\n<b>This</b> <i>is</i> <font color=red face=\"Arial Black\" size=15pt> <underline double>formatted</underline double></font> <s>text</s>.");

			//Set the formatted value
			xlsFirstTable->easy_getCell(1, 0)->setHTMLValue(sFormattedValue); 
			xlsFirstTable->easy_getCell(1, 0)->setDataType(DATATYPE_STRING);
			xlsFirstTable->easy_getCell(1, 0)->setWrap(true); 
			xlsFirstTable->easy_getRowAt(1)->setHeight(250);
			xlsFirstTable->easy_getColumnAt(0)->setWidth(250);



			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial19.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial19.xls");
			
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