<%@ Language=VBScript %>

<%
	'==========================================================================
	' Tutorial 39
	'
	' This tutorial shows how to load a CSV file (we use the file
	' generated in Tutorial 30), modify some data and save it to
	' another file (Tutorial39.xls).
	'==========================================================================
	
	response.write("Tutorial 39<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Read the file
	response.write("Reading file: C:\Samples\Tutorial30.csv<br>")
	if (xls.easy_LoadCSVFile("C:\Samples\Tutorial30.csv")) then
		
		'Set the name of the first worksheet
		xls.easy_getSheetAt(0).setSheetName("First tab")

		'Add a new worksheet and write some data
		xls.easy_addWorksheet_2("Second tab")
		set xlsTable = xls.easy_getSheetAt(1).easy_getExcelTable()
		xlsTable.easy_getCell_2("A1").setValue("Data added by Tutorial39")

		for column=0 to 4
			xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
		
		'Generate the file
		response.write("Writing file: C:\Samples\Tutorial39.xls<br>")
		xls.easy_WriteXLSFile ("C:\Samples\Tutorial39.xls")
		
		'Confirm generation
		if xls.easy_getError() = "" then
			response.write("File successfully created.")
		else
			response.write("Error encountered: " + xls.easy_getError())
		end if
	else
		response.write("Error reading file C:\Samples\Tutorial30.csv")
		response.write(xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
