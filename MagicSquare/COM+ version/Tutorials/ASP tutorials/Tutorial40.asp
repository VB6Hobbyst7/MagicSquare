<%@ Language=VBScript %>

<%
	'==========================================================================
	' Tutorial 40
	'
	' This tutorial shows how to load an HTML file (we use the file
	' generated in Tutorial 31), modify some data and save it to
	' another file (Tutorial40.xls).
	'==========================================================================
	
	response.write("Tutorial 40<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Read the file
	response.write("Reading file: C:\Samples\Tutorial31.html<br>")
	if (xls.easy_LoadHTMLFile_2("C:\Samples\Tutorial31.html")) then
		
		'Set the name of the first worksheet
		xls.easy_getSheetAt(0).setSheetName("First tab")

		'Add a new worksheet and write some data
		xls.easy_addWorksheet_2("Second tab")
		set xlsTable = xls.easy_getSheetAt(1).easy_getExcelTable()
		xlsTable.easy_getCell_2("A1").setValue("Data added by Tutorial40")

		for column=0 to 4
			xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
		
		'Generate the file
		response.write("Writing file: C:\Samples\Tutorial40.xls<br>")
		xls.easy_WriteXLSFile ("C:\Samples\Tutorial40.xls")
		
		'Confirm generation
		if xls.easy_getError() = "" then
			response.write("File successfully created.")
		else
			response.write("Error encountered: " + xls.easy_getError())
		end if
	else
		response.write("Error reading file C:\Samples\Tutorial31.html")
		response.write(xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
