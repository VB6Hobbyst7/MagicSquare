<%@ Language=VBScript %>

<%
	'==========================================================================
	'Tutorial 12
	'
	' This tutorial shows how to create a Microsoft Excel file
	' that has two worksheets. The second one contains a named
	' range.
	'==========================================================================
	
	response.write("Tutorial 12<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
	
	'Get the table of the second worksheet and populate the sheet
	set xlsSecondTab = xls.easy_getSheetAt(1)
	set xlsSecondTable = xlsSecondTab.easy_getExcelTable()
	xlsSecondTable.easy_getCell_2("A1").setValue("Range data 1")
	xlsSecondTable.easy_getCell_2("A2").setValue("Range data 2")
	xlsSecondTable.easy_getCell_2("A3").setValue("Range data 3")
	xlsSecondTable.easy_getCell_2("A4").setValue("Range data 4")

	'Create a named range
	xlsSecondTab.easy_addName_2 "Range", "='Second tab'!$A$1:$A$4"

	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial12.xls<br>")
	xls.easy_WriteXLSFile ("C:\Samples\Tutorial12.xls")
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
