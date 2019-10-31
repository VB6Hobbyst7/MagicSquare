<%@ Language=VBScript %>

<%
	'==========================================================================
	' Tutorial 38
	'
	' This tutorial shows how to load a XLSB file (we use the file
	' generated in Tutorial 29), modify some data and save it to
	' another file (Tutorial38.xlsb).
	'==========================================================================
	
	response.write("Tutorial 38<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Read the file
	response.write("Reading file: C:\Samples\Tutorial29.xlsb<br>")
	if (xls.easy_LoadXLSBFile("C:\Samples\Tutorial29.xlsb")) then

		'Get the table of the second worksheet
		set xlsTable = xls.easy_getSheet("Second tab").easy_getExcelTable()
		
		'Write some data
		xlsTable.easy_getCell_2("A1").setValue("Data added by Tutorial38")

		for column=0 to 4
			xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
		
		'Generate the file
		response.write("Writing file: C:\Samples\Tutorial38.xlsb<br>")
		xls.easy_WriteXLSBFile ("C:\Samples\Tutorial38.xlsb")
		
		'Confirm generation
		if xls.easy_getError() = "" then
			response.write("File successfully created.")
		else
			response.write("Error encountered: " + xls.easy_getError())
		end if
	else
		response.write("Error reading file C:\Samples\Tutorial29.xlsb")
		response.write(xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
