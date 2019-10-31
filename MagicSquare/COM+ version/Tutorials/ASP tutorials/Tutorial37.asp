<%@ Language=VBScript %>

<%
	'==========================================================================
	' Tutorial 37
	'
	' This tutorial shows how to load a XLSX file (we use the file
	' generated in Tutorial 28), modify some data and save it to
	' another file (Tutorial37.xlsx).
	'==========================================================================
	
	response.write("Tutorial 37<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Read the file
	response.write("Reading file: C:\Samples\Tutorial28.xlsx<br>")
	if (xls.easy_LoadXLSXFile("C:\Samples\Tutorial28.xlsx")) then

		'Get the table of the second worksheet
		set xlsTable = xls.easy_getSheet("Second tab").easy_getExcelTable()
		
		'Write some data
		xlsTable.easy_getCell_2("A1").setValue("Data added by Tutorial37")

		for column=0 to 4
			xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
		
		'Generate the file
		response.write("Writing file: C:\Samples\Tutorial37.xlsx<br>")
		xls.easy_WriteXLSXFile ("C:\Samples\Tutorial37.xlsx")
		
		'Confirm generation
		if xls.easy_getError() = "" then
			response.write("File successfully created.")
		else
			response.write("Error encountered: " + xls.easy_getError())
		end if
	else
		response.write("Error reading file C:\Samples\Tutorial28.xlsx")
		response.write(xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
