<%@ Language=VBScript %>

<!-- #INCLUDE FILE="DataType.inc" -->
<%
	'==========================================================================
	'Tutorial 30
	'
	' This tutorial shows how to export a CSV file.
	'==========================================================================
	
	response.write("Tutorial 30<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Create the worksheet
	xls.easy_addWorksheet_2("First tab")
	
	'Get the table of the first worksheet
	Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
	
	'Add the cells for header
	for column = 0 to 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (column + 1))
		xlsFirstTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	next

	'Add the cells for data
	for row = 0 to 99
		for column = 0 to 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		next
	next


	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial30.csv<br>")
	xls.easy_WriteCSVFile "C:\Samples\Tutorial30.csv", "First tab"
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
