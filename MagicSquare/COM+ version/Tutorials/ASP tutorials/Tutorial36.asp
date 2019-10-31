<%@ Language=VBScript %>

<%
	'==========================================================================
	' Tutorial 36
	'
	' This tutorial shows how to load an excel file (we use the one
	' generated in Tutorial 09), modify some data and save it to
	' another file (Tutorial36.xls).
	'==========================================================================
	
	response.write("Tutorial 36<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Read the file
	response.write("Reading file: C:\Samples\Tutorial09.xls<br>")
	if (xls.easy_LoadXLSFile("C:\Samples\Tutorial09.xls")) then
		'Get the table of the second worksheet
		set xlsSecondTable = xls.easy_getSheet("Second tab").easy_getExcelTable()
		
		'Write some data
		xlsSecondTable.easy_getCell_2("A1").setValue("Data added by Tutorial36")

		for column=0 to 4
			xlsSecondTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
		
		'Generate the file
		response.write("Writing file: C:\Samples\Tutorial36.xls<br>")
		xls.easy_WriteXLSFile ("C:\Samples\Tutorial36.xls")
		
		'Confirm generation
		if xls.easy_getError() = "" then
			response.write("File successfully created.")
		else
			response.write("Error encountered: " + xls.easy_getError())
		end if
	else
		response.write("Error reading file C:\Samples\Tutorial09.xls ")
		response.write(xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
