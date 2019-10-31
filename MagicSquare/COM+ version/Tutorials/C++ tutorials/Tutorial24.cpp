/* ----------------------------------------------------------------
 * Tutorial 24
 * 
 * This tutorial shows how to create and insert a chart in a worksheet. 
 * ----------------------------------------------------------------- */

#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 24\n----------\n");

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

			//Add one worksheet
			xls->easy_addWorksheet_2("SourceData");
			
			// ----------------------------------------------------------------------
			//Insert values
			EasyXLS::IExcelWorksheetPtr xlsFirstTab = (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheet("SourceData");	
			EasyXLS::IExcelTablePtr xlsTable1 = xlsFirstTab->easy_getExcelTable();

			xlsTable1->easy_getCell(0, 0)->setValue("Show Date");
			xlsTable1->easy_getCell(0, 1)->setValue("Available Places");
			xlsTable1->easy_getCell(0, 2)->setValue("Available Tickets");
			xlsTable1->easy_getCell(0, 3)->setValue("Sold Tickets");

			xlsTable1->easy_getCell(1, 0)->setValue("03/13/2005 00:00:00");
			xlsTable1->easy_getCell(1, 0)->setFormat(FORMAT_FORMAT_DATE);
			xlsTable1->easy_getCell(2, 0)->setValue("03/14/2005 00:00:00");
			xlsTable1->easy_getCell(2, 0)->setFormat(FORMAT_FORMAT_DATE);
			xlsTable1->easy_getCell(3, 0)->setValue("03/15/2005 00:00:00");
			xlsTable1->easy_getCell(3, 0)->setFormat(FORMAT_FORMAT_DATE);
			xlsTable1->easy_getCell(4, 0)->setValue("03/16/2005 00:00:00");
			xlsTable1->easy_getCell(4, 0)->setFormat(FORMAT_FORMAT_DATE);

			xlsTable1->easy_getCell(1, 1)->setValue("10000");
			xlsTable1->easy_getCell(2, 1)->setValue("5000");
			xlsTable1->easy_getCell(3, 1)->setValue("8500");
			xlsTable1->easy_getCell(4, 1)->setValue("1000");

			xlsTable1->easy_getCell(1, 2)->setValue("8000");
			xlsTable1->easy_getCell(2, 2)->setValue("4000");
			xlsTable1->easy_getCell(3, 2)->setValue("6000");
			xlsTable1->easy_getCell(4, 2)->setValue("1000");

			xlsTable1->easy_getCell(1, 3)->setValue("920");
			xlsTable1->easy_getCell(2, 3)->setValue("1005");
			xlsTable1->easy_getCell(3, 3)->setValue("342");
			xlsTable1->easy_getCell(4, 3)->setValue("967");

			xlsTable1->easy_getColumnAt(0)->setWidth(100);
			xlsTable1->easy_getColumnAt(1)->setWidth(100);
			xlsTable1->easy_getColumnAt(2)->setWidth(100);
			xlsTable1->easy_getColumnAt(3)->setWidth(100);

			//--------------------------------------------------------------------------
			
			//Create the chart.
			EasyXLS::IExcelChartPtr xlsChart;
			CoCreateInstance(__uuidof(EasyXLS::ExcelChart), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelChart), (void**) &xlsChart) ;
			xlsChart->setLeftUpperCorner_2("A10");
			xlsChart->setSize(600, 300);
				
			xlsChart->easy_addSeries_2("=SourceData!$B$1", "=SourceData!$B$2:$B$5");
			xlsChart->easy_addSeries_2("=SourceData!$C$1", "=SourceData!$C$2:$C$5");
			xlsChart->easy_addSeries_2("=SourceData!$D$1", "=SourceData!$D$2:$D$5");
			xlsChart->easy_setCategoryXAxisLabels("=SourceData!$A$2:$A$5");
		
			//Add the chart to the first worksheet.	
			EasyXLS::IExcelWorksheetPtr xlsWorksheet = (EasyXLS::IExcelWorksheetPtr)xls->easy_getSheet("SourceData");
			xlsWorksheet->easy_addChart(xlsChart);
			
			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial24.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial24.xls");
			
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