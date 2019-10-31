<%@ Language=VBScript %>

<!-- #INCLUDE FILE="Format.inc" -->
<!-- #INCLUDE FILE="Chart.inc" -->

<%
	'==========================================================================
	' Tutorial 21
	'
	' This tutorial helps you to create a simple chart with a legend.
	'==========================================================================
	
	response.write("Tutorial 21<br>")
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
	
	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial21.xls<br>")
	xls.easy_WriteXLSFile ("C:\Samples\Tutorial21.xls")
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
