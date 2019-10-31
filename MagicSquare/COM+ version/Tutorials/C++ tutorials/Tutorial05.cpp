/* ----------------------------------------------------------------
 * Tutorial 05
 * 
 * This tutorial shows how to create a Microsoft Excel file
 * that has two worksheets. The first one is full with data
 * and the cells are formatted.
 * ----------------------------------------------------------------- */


#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 05\n----------\n");

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
			
			//Create the style for the header
			EasyXLS::IExcelStylePtr xlsStyleHeader;
			hr = CoCreateInstance(__uuidof(EasyXLS::ExcelStyle),
			NULL,
			CLSCTX_ALL,
			__uuidof(EasyXLS::IExcelStyle),
			(void**) &xlsStyleHeader) ;

			xlsStyleHeader->setFont("Verdana");
			xlsStyleHeader->setFontSize(8);
			xlsStyleHeader->setItalic(true);
			xlsStyleHeader->setBold(true);
			xlsStyleHeader->setForeground(COLOR_YELLOW);
			xlsStyleHeader->setBackground(COLOR_BLACK);
			xlsStyleHeader->setBorderColors (COLOR_GRAY, COLOR_GRAY, COLOR_GRAY, COLOR_GRAY);
			xlsStyleHeader->setBorderStyles (BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM);
			xlsStyleHeader->setHorizontalAlignment(ALIGNMENT_ALIGNMENT_CENTER);
			xlsStyleHeader->setVerticalAlignment(ALIGNMENT_ALIGNMENT_BOTTOM);
			xlsStyleHeader->setWrap(true);
			xlsStyleHeader->setDataType(DATATYPE_STRING);

			//Add the cells for header
			char* cellValue = (char*)malloc(11*sizeof(char));
			char*  columnNumber = (char*)malloc(sizeof(char));
			for (int column=0; column<5; column++)
			{
				strcpy_s(cellValue, 8, "Column ");			
				_itoa_s(column+ 1, columnNumber, 2, 10);
				strcat_s(cellValue, 10, columnNumber);
				xlsFirstTable->easy_getCell(0,column)->setValue(cellValue); 
				xlsFirstTable->easy_getCell(0,column)->setStyle(xlsStyleHeader); 
			}
			xlsFirstTable->easy_getRowAt(0)->setHeight(30);

			//Create a style for cells
			EasyXLS::IExcelStylePtr xlsStyleData;
			hr = CoCreateInstance(__uuidof(EasyXLS::ExcelStyle),
			NULL,
			CLSCTX_ALL,
			__uuidof(EasyXLS::IExcelStyle),
			(void**) &xlsStyleData) ;

			xlsStyleData->setHorizontalAlignment(ALIGNMENT_ALIGNMENT_LEFT);
			xlsStyleData->setForeground(COLOR_DARKGRAY);
			xlsStyleData->setWrap(false);
			xlsStyleData->setDataType(DATATYPE_STRING);

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

					xlsFirstTable->easy_getCell(row+1,column)->setValue(cellValue); 
					xlsFirstTable->easy_getCell(row+1,column)->setStyle(xlsStyleData);
				}
			}

			//Set column widths
			xlsFirstTable->setColumnWidth_2(0, 70);
			xlsFirstTable->setColumnWidth_2(1, 100);
			xlsFirstTable->setColumnWidth_2(2, 70);
			xlsFirstTable->setColumnWidth_2(3, 100);
			xlsFirstTable->setColumnWidth_2(4, 70);			

		
			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial05.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial05.xls");
			
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