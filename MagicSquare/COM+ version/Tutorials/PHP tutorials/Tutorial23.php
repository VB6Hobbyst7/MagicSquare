<?php
	/*==========================================================================
	 | Tutorial 23
	 |
	 | This tutorial shows how to modify different properties          |
	 | related to the chart.
	  ==========================================================================*/
	
	//Include Files
	include("Format.inc");
	include("Color.inc");
	include("Chart.inc");
	include("LineStyleFormat.inc");
	include("ShadowFormat.inc");


	header("Content-Type: text/html");
	
	echo "Tutorial 23<br>";
	echo "----------<br>";
	



	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Add one worksheet
	$xls->easy_addWorksheet_2("SourceData");
	
	// ----------------------------------------------------------------------
	//Insert values	
	$xlsTable1 = $xls->easy_getSheet("SourceData")->easy_getExcelTable();

	$xlsTable1->easy_getCell(0, 0)->setValue("Show Date");
	$xlsTable1->easy_getCell(0, 1)->setValue("Available Places");
	$xlsTable1->easy_getCell(0, 2)->setValue("Available Tickets");
	$xlsTable1->easy_getCell(0, 3)->setValue("Sold Tickets");

	$xlsTable1->easy_getCell(1, 0)->setValue("03/13/2005 00:00:00");
	$xlsTable1->easy_getCell(1, 0)->setFormat($FORMAT_FORMAT_DATE);
	$xlsTable1->easy_getCell(2, 0)->setValue("03/14/2005 00:00:00");
	$xlsTable1->easy_getCell(2, 0)->setFormat($FORMAT_FORMAT_DATE);
	$xlsTable1->easy_getCell(3, 0)->setValue("03/15/2005 00:00:00");
	$xlsTable1->easy_getCell(3, 0)->setFormat($FORMAT_FORMAT_DATE);
	$xlsTable1->easy_getCell(4, 0)->setValue("03/16/2005 00:00:00");
	$xlsTable1->easy_getCell(4, 0)->setFormat($FORMAT_FORMAT_DATE);

	$xlsTable1->easy_getCell(1, 1)->setValue("10000");
	$xlsTable1->easy_getCell(2, 1)->setValue("5000");
	$xlsTable1->easy_getCell(3, 1)->setValue("8500");
	$xlsTable1->easy_getCell(4, 1)->setValue("1000");

	$xlsTable1->easy_getCell(1, 2)->setValue("8000");
	$xlsTable1->easy_getCell(2, 2)->setValue("4000");
	$xlsTable1->easy_getCell(3, 2)->setValue("6000");
	$xlsTable1->easy_getCell(4, 2)->setValue("1000");

	$xlsTable1->easy_getCell(1, 3)->setValue("920");
	$xlsTable1->easy_getCell(2, 3)->setValue("1005");
	$xlsTable1->easy_getCell(3, 3)->setValue("342");
	$xlsTable1->easy_getCell(4, 3)->setValue("967");

	$xlsTable1->easy_getColumnAt(0)->setWidth(100);
	$xlsTable1->easy_getColumnAt(1)->setWidth(100);
	$xlsTable1->easy_getColumnAt(2)->setWidth(100);
	$xlsTable1->easy_getColumnAt(3)->setWidth(100);

	//--------------------------------------------------------------------------

	//Add the chart
	echo "";
	$xls->easy_addChart_5("Chart", "=SourceData!\$A$1:\$D$5", $CHART_SERIES_IN_COLUMNS);
	
	//Get the previously added chart
	$xlsChartSheet = $xls->easy_getSheetAt(1);
	$xlsChart = $xlsChartSheet->easy_getExcelChart();

	//Modifying chart type
	$xlsChart->easy_setChartType($CHART_CHART_TYPE_CYLINDER_COLUMN);

	//Modifying chart area properties
	$xlsChartArea = $xlsChart->easy_getChartArea();
	$xlsChartArea->getLineColorFormat()->setLineColor((int)$COLOR_DARKGRAY);
	$xlsChartArea->getLineStyleFormat()->setDashType($LINESTYLEFORMAT_DASH_TYPE_SOLID);
	$xlsChartArea->getLineStyleFormat()->setWidth(0.25);
	
	//Modifying chart plot area properties
	$xlsPlotArea = $xlsChart->easy_getPlotArea();
	$xlsPlotArea->getLineColorFormat()->setLineColor((int)$COLOR_DARKGRAY);
	$xlsPlotArea->getLineStyleFormat()->setDashType($LINESTYLEFORMAT_DASH_TYPE_SOLID);
	$xlsPlotArea->getLineStyleFormat()->setWidth(0.25);

	//Modifying legend property
	$xlsChatLegend = $xlsChart->easy_getLegend();
	$xlsChatLegend->getFillFormat()->setBackground((int)$COLOR_LAVENDERBLUSH);
	$xlsChatLegend->getFontFormat()->setForeground((int)$COLOR_BLUE);
	$xlsChatLegend->getFontFormat()->setItalic(true);
	$xlsChatLegend->setKeysArrangementDirection($CHART_KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL);
	$xlsChatLegend->setPlacement($CHART_LEGEND_CORNER);
	$xlsChatLegend->getShadowFormat()->setShadow($SHADOWFORMAT_OFFSET_DIAGONAL_BOTTOM_RIGHT);

	//Modifying X axis properties
	$xlsXAxis = $xlsChart->easy_getCategoryXAxis();
	$xlsXAxis->getLineColorFormat()->setLineColor((int)$COLOR_STEELBLUE);
	$xlsXAxis->getLineStyleFormat()->setDashType($LINESTYLEFORMAT_DASH_TYPE_DASH_DOT);
	$xlsXAxis->getLineStyleFormat()->setWidth(0.25);
	$xlsXAxis->getFontFormat()->setForeground((int)$COLOR_RED);

	//Modifying Y axis properties
	$xlsYAxis = $xlsChart->easy_getValueYAxis();
	$xlsYAxis->getLineColorFormat()->setLineColor((int)$COLOR_STEELBLUE);
	$xlsYAxis->getLineStyleFormat()->setDashType($LINESTYLEFORMAT_DASH_TYPE_LONG_DASH);
	$xlsYAxis->getLineStyleFormat()->setWidth(0.25);
	$xlsYAxis->getFontFormat()->setForeground((int)$COLOR_BLUE);

	//Modifying series properties
	$xlsChart->easy_getSeriesAt(0)->getFillFormat()->setBackground((int)$COLOR_ROYALBLUE);
	$xlsChart->easy_getSeriesAt(1)->getFillFormat()->setBackground((int)$COLOR_YELLOW);
	$xlsChart->easy_getSeriesAt(2)->getFillFormat()->setBackground((int)$COLOR_LIGHTGREEN);

	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial23.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial23.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>