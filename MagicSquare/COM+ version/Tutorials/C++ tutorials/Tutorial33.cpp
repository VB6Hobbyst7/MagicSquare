/* -----------------------------------------------------------------
 * Tutorial 33
 * 
 * This tutorial shows how to set the properties of the document.
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 33\n----------\n");

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
			
			//Set the 'Subject' property
			xls->getSummaryInformation()->setSubject("This is the subject");

			//Set the 'Manager' property
			xls->getDocumentSummaryInformation()->setManager("This is the manager");

			//Set a custom property
			xls->getDocumentSummaryInformation()->setCustomProperty("PropertyName", FILEPROPERTY_VT_NUMBER, "4");
			
			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial33.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial33.xls");
			
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
