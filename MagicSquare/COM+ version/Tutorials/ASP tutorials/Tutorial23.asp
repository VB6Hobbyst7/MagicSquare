<%@ Language=VBScript %>

<!-- #INCLUDE FILE="Chart.inc" -->
<!-- #INCLUDE FILE="Color.inc" -->
<!-- #INCLUDE FILE="Format.inc" -->
<!-- #INCLUDE FILE="LineStyleFormat.inc" -->
<!-- #INCLUDE FILE="ShadowFormat.inc" -->
<%
	'==========================================================================
	' Tutorial 23
	'
	' This tutorial shows how to modify different properties
	' related to the chart.
	'==========================================================================
	
	response.write("Tutorial 23<br>")
	response.write("----------<br>")



	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Add one worksheet
	xls.easy_addWorksheet_2("SourceData")
	
	'----------------------------------------------------------------------
	'Insert values
	Set xlsTable1 = xls.easy_getSheet("SourceData").easy_getExcelTable()

	xlsTable1.easy_getCell(0, 0).setValue("Show Date")
	xlsTable1.easy_getCell(0, 1).setValue("Available Places")
	xlsTable1.easy_getCell(0, 2).setValue("Available Tickets")
	xlsTable1.easy_getCell(0, 3).setValue("Sold Tickets")

	xlsTable1.easy_getCell(1, 0).setValue("03/13/2005 00:00:00")
	xlsTable1.easy_getCell(1, 0).setFormat(FORMAT_FORMAT_DATE)
	xlsTable1.easy_getCell(2, 0).setValue("03/14/2005 00:00:00")
	xlsTable1.easy_getCell(2, 0).setFormat(FORMAT_FORMAT_DATE)
	xlsTable1.easy_getCell(3, 0).setValue("03/15/2005 00:00:00")
	xlsTable1.easy_getCell(3, 0).setFormat(FORMAT_FORMAT_DATE)
	xlsTable1.easy_getCell(4, 0).setValue("03/16/2005 00:00:00")
	xlsTable1.easy_getCell(4, 0).setFormat(FORMAT_FORMAT_DATE)
	
	xlsTable1.easy_getCell(1, 1).setValue("10000")
	xlsTable1.easy_getCell(2, 1).setValue("5000")
	xlsTable1.easy_getCell(3, 1).setValue("8500")
	xlsTable1.easy_getCell(4, 1).setValue("1000")

	xlsTable1.easy_getCell(1, 2).setValue("8000")
	xlsTable1.easy_getCell(2, 2).setValue("4000")
	xlsTable1.easy_getCell(3, 2).setValue("6000")
	xlsTable1.easy_getCell(4, 2).setValue("1000")

	xlsTable1.easy_getCell(1, 3).setValue("920")
	xlsTable1.easy_getCell(2, 3).setValue("1005")
	xlsTable1.easy_getCell(3, 3).setValue("342")
	xlsTable1.easy_getCell(4, 3).setValue("967")

	xlsTable1.easy_getColumnAt(0).setWidth(100)
	xlsTable1.easy_getColumnAt(1).setWidth(100)
	xlsTable1.easy_getColumnAt(2).setWidth(100)
	xlsTable1.easy_getColumnAt(3).setWidth(100)

	'--------------------------------------------------------------------------
	
	'Add the chart
	xls.easy_addChart_5 "Chart", "=SourceData!$A$1:$D$5", CHART_SERIES_IN_COLUMNS

	'Get the previously added chart
	Set xlsChartSheet = xls.easy_getSheetAt(1)
	Set xlsChart = xlsChartSheet.easy_getExcelChart()

	'Modifying chart type
	xlsChart.easy_setChartType(CHART_CHART_TYPE_CYLINDER_COLUMN)

	'Modifying chart area properties
	Set xlsChartArea = xlsChart.easy_getChartArea()
	xlsChartArea.getLineColorFormat().setLineColor(CLng(COLOR_DARKGRAY))
	xlsChartArea.getLineStyleFormat().setDashType(LINESTYLEFORMAT_DASH_TYPE_SOLID)
	xlsChartArea.getLineStyleFormat().setWidth(0.25)
	
	'Modifying chart plot area properties
	Set xlsPlotArea = xlsChart.easy_getPlotArea()
	xlsPlotArea.getLineColorFormat().setLineColor(CLng(COLOR_DARKGRAY))
	xlsPlotArea.getLineStyleFormat().setDashType(LINESTYLEFORMAT_DASH_TYPE_SOLID)
	xlsPlotArea.getLineStyleFormat().setWidth(0.25)

	'Modifying legend property
	Set xlsChartLegend = xlsChart.easy_getLegend()
	xlsChartLegend.getFillFormat().setBackground(CLng(COLOR_LAVENDERBLUSH))
	xlsChartLegend.getFontFormat().setForeground(CLng(COLOR_BLUE))
	xlsChartLegend.getFontFormat().setItalic(True)
	xlsChartLegend.setKeysArrangementDirection(CHART_KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL)
	xlsChartLegend.setPlacement(CHART_LEGEND_CORNER)
	xlsChartLegend.getShadowFormat().setShadow(SHADOWFORMAT_OFFSET_DIAGONAL_BOTTOM_RIGHT)

	'Modifying X axis properties
	Set xlsXAxis = xlsChart.easy_getCategoryXAxis()
	xlsXAxis.getLineColorFormat().setLineColor(CLng(COLOR_STEELBLUE))
	xlsXAxis.getLineStyleFormat().setDashType(LINESTYLEFORMAT_DASH_TYPE_DASH_DOT)
	xlsXAxis.getLineStyleFormat().setWidth(0.25)
	xlsXAxis.getFontFormat().setForeground(CLng(COLOR_RED))

	'Modifying Y axis properties
	Set xlsYAxis = xlsChart.easy_getValueYAxis()
	xlsYAxis.getLineColorFormat().setLineColor(CLng(COLOR_STEELBLUE))
	xlsYAxis.getLineStyleFormat().setDashType(LINESTYLEFORMAT_DASH_TYPE_LONG_DASH)
	xlsYAxis.getLineStyleFormat().setWidth(0.25)
	xlsYAxis.getFontFormat().setForeground(CLng(COLOR_BLUE))

	'Modifying series properties
	xlsChart.easy_getSeriesAt(0).getFillFormat().setBackground(CLng(COLOR_ROYALBLUE))
	xlsChart.easy_getSeriesAt(1).getFillFormat().setBackground(CLng(COLOR_YELLOW))
	xlsChart.easy_getSeriesAt(2).getFillFormat().setBackground(CLng(COLOR_LIGHTGREEN))

	
	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial23.xls<br>")
	xls.easy_WriteXLSFile ("C:\Samples\Tutorial23.xls")
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
