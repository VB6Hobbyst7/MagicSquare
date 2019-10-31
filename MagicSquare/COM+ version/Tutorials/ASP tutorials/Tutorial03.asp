<%@ Language=VBScript %>

<%
	'==========================================================================
	' Tutorial 03
	'
	' This tutorial shows how to create a Microsoft Excel file
	' that has two worksheets.
	'==========================================================================

	response.write("Tutorial 03<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")

	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial03.xls<br>")
	xls.easy_WriteXLSFile ("C:\Samples\Tutorial03.xls")
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
