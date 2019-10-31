<%@ Language=VBScript %>

<!-- #INCLUDE FILE="Format.inc" -->

<%
	'==========================================================================
	' Tutorial 20
	'
	' This tutorial shows how to create a Microsoft Excel file 
    ' that has AutoFilter.
	'==========================================================================
	
	response.write("Tutorial 20<br>")
	response.write("----------<br>")

	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Add one worksheet
	xls.easy_addWorksheet_2("Sheet1")
	
	'Get the table of the first worksheet
	set xlsTab = xls.easy_getSheet("Sheet1")
	set xlsTable = xlsTab.easy_getExcelTable()
	
	'Add the cells for header
	for column = 0 to 4
		xlsTable.easy_getCell(0,column).setValue("Column " & (column + 1))
		xlsTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	next
	
	'Add the cells for data
	for row = 0 to 99
		for column = 0 to 4
			xlsTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		next
	next

	'Add AutoFilter
	set xlsFilter = xlsTab.easy_getFilter()
	xlsFilter.setAutoFilter_2("A1:E1")
	
	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial20.xls<br>")
	xls.easy_WriteXLSFile ("C:\Samples\Tutorial20.xls")
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
