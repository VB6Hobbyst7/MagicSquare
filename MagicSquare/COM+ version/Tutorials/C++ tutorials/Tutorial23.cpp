/* -----------------------------------------------------------------
 * Tutorial 23
 *
 * This tutorial shows how to modify different properties
 * related to the chart.
 * ----------------------------------------------------------------- */



#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>


int _tmain(int argc, _TCHAR* argv[])
{
	printf("Tutorial 23\n----------\n");

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

			//Add the chart
			xls->easy_addChart_5("Chart", "=SourceData!$A$1:$D$5", CHART_SERIES_IN_COLUMNS);

			//Get the previously added chart
			EasyXLS::IExcelChartSheetPtr xlsChartSheet = (EasyXLS::IExcelChartSheetPtr)xls->easy_getSheetAt(1);	
			EasyXLS::IExcelChartPtr xlsChart = xlsChartSheet->easy_getExcelChart();

			//Modifying chart type
			xlsChart->easy_setChartType(CHART_CHART_TYPE_CYLINDER_COLUMN);

			//Modifying chart area properties
			EasyXLS::IExcelChartAreaPtr xlsChartArea = xlsChart->easy_getChartArea();
			xlsChartArea->getLineColorFormat()->setLineColor(COLOR_DARKGRAY);
			xlsChartArea->getLineStyleFormat()->setDashType(LINESTYLEFORMAT_DASH_TYPE_SOLID);
			xlsChartArea->getLineStyleFormat()->setWidth(0.25f);
			
			//Modifying chart plot area properties
			EasyXLS::IExcelPlotAreaPtr xlsPlotArea = xlsChart->easy_getPlotArea();
			xlsPlotArea->getLineColorFormat()->setLineColor(COLOR_DARKGRAY);
			xlsPlotArea->getLineStyleFormat()->setDashType(LINESTYLEFORMAT_DASH_TYPE_SOLID);
			xlsPlotArea->getLineStyleFormat()->setWidth(0.25f);

			//Modifying legend property
			EasyXLS::IExcelChartLegendPtr xlsChartLegend = xlsChart->easy_getLegend();
			xlsChartLegend->getFillFormat()->setBackground(COLOR_LAVENDERBLUSH);
			xlsChartLegend->getFontFormat()->setForeground(COLOR_BLUE);
			xlsChartLegend->getFontFormat()->setItalic(true);
			xlsChartLegend->setKeysArrangementDirection(CHART_KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL);
			xlsChartLegend->setPlacement(CHART_LEGEND_CORNER);
			xlsChartLegend->getShadowFormat()->setShadow(SHADOWFORMAT_OFFSET_DIAGONAL_BOTTOM_RIGHT);

			//Modifying X axis properties
			EasyXLS::IExcelAxisPtr xlsXAxis = xlsChart->easy_getCategoryXAxis();
			xlsXAxis->getLineColorFormat()->setLineColor(COLOR_STEELBLUE);
			xlsXAxis->getLineStyleFormat()->setDashType(LINESTYLEFORMAT_DASH_TYPE_DASH_DOT);
			xlsXAxis->getLineStyleFormat()->setWidth(0.25f);
			xlsXAxis->getFontFormat()->setForeground(COLOR_RED);

			//Modifying Y axis properties
			EasyXLS::IExcelAxisPtr xlsYAxis = xlsChart->easy_getValueYAxis();
			xlsYAxis ->getLineColorFormat()->setLineColor(COLOR_STEELBLUE);
			xlsYAxis ->getLineStyleFormat()->setDashType(LINESTYLEFORMAT_DASH_TYPE_LONG_DASH);
			xlsYAxis ->getLineStyleFormat()->setWidth(0.25f);
			xlsYAxis ->getFontFormat()->setForeground(COLOR_BLUE);

			//Modifying series properties
			xlsChart->easy_getSeriesAt(0)->getFillFormat()->setBackground(COLOR_ROYALBLUE);
			xlsChart->easy_getSeriesAt(1)->getFillFormat()->setBackground(COLOR_YELLOW);
			xlsChart->easy_getSeriesAt(2)->getFillFormat()->setBackground(COLOR_LIGHTGREEN);


			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial23.xls.");
			xls->easy_WriteXLSFile("C:\\Samples\\Tutorial23.xls");
			
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