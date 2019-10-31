<%@ Language=VBScript %>

<%
	'==========================================================================
	' Tutorial 27
	'
	' This tutorial shows how to encrypt and set the password required for opening a document
	'==========================================================================

	response.write("Tutorial 27<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")

	'Set the password required for opening the document
	xls.easy_getOptions().setPasswordToOpen("password")

	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial27.xls<br>")
	xls.easy_WriteXLSFile ("C:\Samples\Tutorial27.xls")
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
