<%@ Language=VBScript %>

<!-- #INCLUDE FILE="DataType.inc" -->
<!-- #INCLUDE FILE="ConditionalFormatting.inc" -->
<!-- #INCLUDE FILE="Color.inc" -->
<%
	'==========================================================================
	' Tutorial 14
	'
	' This tutorial shows how to create conditional formatting ranges.
	'==========================================================================
	
	response.write("Tutorial 14<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Add the worksheet
	xls.easy_addWorksheet_2("Sheet1")
	
	'Insert values
	set xlsTab = xls.easy_getSheet("Sheet1")
	set xlsTable = xlsTab.easy_getExcelTable()

	for i=0 to 5
		for j=0 to 3
			if ( (i<2) and (j<2) ) then
				xlsTable.easy_getCell(i, j).setValue("12")
			else
				if ( (j=2) and (i<2) ) then
					xlsTable.easy_getCell(i, j).setValue("1000")
				else
					xlsTable.easy_getCell(i, j).setValue("9")
				end if
			end if
			xlsTable.easy_getCell(i, j).setDataType(DATATYPE_NUMERIC)
		next
	next
	
	'Set a conditional formatting
	xlsTab.easy_addConditionalFormatting_5 "A1:C3", CONDITIONALFORMATTING_OPERATOR_BETWEEN, "=9", "=11", true, true, Clng(COLOR_RED)

	'Set a conditional formatting
	xlsTab.easy_addConditionalFormatting_9 "A6:C6", CONDITIONALFORMATTING_OPERATOR_BETWEEN, "=COS(PI())+2", "", Clng(COLOR_BISQUE)
	xlsTab.easy_getConditionalFormattingAt_2("A6:C6").getConditionAt(0).setConditionType(CONDITIONALFORMATTING_CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA)
	
	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial14.xls<br>")
	xls.easy_WriteXLSFile ("C:\Samples\Tutorial14.xls")
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>