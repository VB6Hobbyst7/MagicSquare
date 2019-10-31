<%@ Language=VBScript %>

<!-- #INCLUDE FILE="DataType.inc" -->
<!-- #INCLUDE FILE="Styles.inc" -->
<%
	'==========================================================================
	'Tutorial 32
	'
	' This tutorial shows how to export an XML file.
	'==========================================================================
	
	response.write("Tutorial 32<br>")
	response.write("----------<br>")



	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
	
	'Get the table of the first worksheet
	Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
	
	'Add the cells for header
	for column = 0 to 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (column + 1))
		xlsFirstTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	next
	xlsFirstTable.easy_getRowAt(0).setHeight(30)

	'Add the cells for data
	for row = 0 to 99
		for column = 0 to 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		next
	next


	'Create an instance of the object used to format the cells
	Dim xlsAutoFormat 
	set xlsAutoFormat = Server.CreateObject("EasyXLS.ExcelAutoFormat")
	xlsAutoFormat.InitAs(AUTOFORMAT_EASYXLS1)

	'Apply the predefined format to the cells.
	xlsFirstTable.easy_setRangeAutoFormat_2 "A1:E101", xlsAutoFormat


	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial32.xml<br>")
	xls.easy_WriteXMLFile_2 ("C:\Samples\Tutorial32.xml")
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
