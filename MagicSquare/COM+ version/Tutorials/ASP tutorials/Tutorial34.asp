<%@ Language=VBScript %>

<%
	'==========================================================================
	' Tutorial 34
	'
	' This tutorial shows how to read values from the active sheet
	' of an excel file (the file generated in Tutorial 09).
	'==========================================================================
	
	response.write("Tutorial 34<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Read the file
	response.write("Reading file: C:\Samples\Tutorial09.xls<br><br>")
	Set rows = xls.easy_ReadXLSActiveSheet_AsList("C:\Samples\Tutorial09.xls")

    if xls.easy_getError() = "" then
		'Display the values
		for rowIndex = 0 to rows.size() - 1
			Set row = rows.elementAt(rowIndex)
			for cellIndex = 0 to row.size - 1
				response.write("At row " & (rowIndex + 1) & ", column " & (cellIndex + 1) & " the value is '" & row.elementAt(cellIndex) & "'<br>")
			next
		next
    else
		response.Write("Error reading file C:\Samples\Tutorial09.xls " & xls.easy_getError())
    end if

	'Dispose memory
	xls.Dispose
%>
