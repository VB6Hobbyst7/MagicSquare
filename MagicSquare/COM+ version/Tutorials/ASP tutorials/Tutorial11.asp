<%@ Language=VBScript %>

<%
	'==========================================================================
	'Tutorial 11
	'
	' This tutorial shows how to create a Microsoft Excel file
	' that has a formula.
	'==========================================================================
	
	response.write("Tutorial 11<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Add one worksheet
	xls.easy_addWorksheet_2("Formula")
	
	'Get the table, populate the sheet and set a formula
	Set xlsTable = xls.easy_getSheet("Formula").easy_getExcelTable()
	xlsTable.easy_getCell_2("A1").setValue("1")
	xlsTable.easy_getCell_2("A2").setValue("2")
	xlsTable.easy_getCell_2("A3").setValue("3")
	xlsTable.easy_getCell_2("A4").setValue("4")
	xlsTable.easy_getCell_2("A6").setValue("=SUM(A1:A4)")

	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial11.xls<br>")
	xls.easy_WriteXLSFile ("C:\Samples\Tutorial11.xls")
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
