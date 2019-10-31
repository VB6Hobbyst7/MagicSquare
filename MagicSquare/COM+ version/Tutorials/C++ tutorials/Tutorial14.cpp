/* -----------------------------------------------------------------
 * Tutorial 14
 * 
 * This tutorial shows how to create conditional formatting ranges.
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 14\n----------\n");

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

			//Add a worksheet
			xls->easy_addWorksheet_2("Sheet1");

			//Get the table of the first sheet
			EasyXLS::IExcelWorksheetPtr xlsTab = (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheet("Sheet1");	
			EasyXLS::IExcelTablePtr xlsTable = xlsTab->easy_getExcelTable();

			//Insert values
			for (int i=0; i<6; i++)
			{
				for (int j=0; j<4; j++)
				{
					if((i<2)&&(j<2))
						xlsTable->easy_getCell(i, j)->setValue("12");
					else
						if((j==2)&&(i<2))
							xlsTable->easy_getCell(i, j)->setValue("1000");
						else
							xlsTable->easy_getCell(i, j)->setValue("9");
					xlsTable->easy_getCell(i, j)->setDataType(DATATYPE_NUMERIC) ;
				}
			}

			//Set a conditional formatting
			xlsTab->easy_addConditionalFormatting_5("A1:C3", CONDITIONALFORMATTING_OPERATOR_BETWEEN, 
				"=9", "=11", true, true, COLOR_RED);

			//Set a conditional formatting
			xlsTab->easy_addConditionalFormatting_9("A6:C6", CONDITIONALFORMATTING_OPERATOR_BETWEEN, "=COS(PI())+2", "", COLOR_BISQUE);
			xlsTab->easy_getConditionalFormattingAt_2("A6:C6")->getConditionAt(0)->setConditionType(CONDITIONALFORMATTING_CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA);


			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial14.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial14.xls");
			
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