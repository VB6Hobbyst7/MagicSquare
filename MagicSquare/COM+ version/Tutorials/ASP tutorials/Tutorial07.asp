<%@ Language=VBScript %>

<!-- #INCLUDE FILE="Alignment.inc" -->
<!-- #INCLUDE FILE="Border.inc" -->
<!-- #INCLUDE FILE="DataType.inc" -->
<!-- #INCLUDE FILE="Color.inc" -->
<%
	'==========================================================================
	'Tutorial 07
	'
	' This tutorial shows how to create a Microsoft Excel file
	' that has two worksheets. The first one is full with data
	' and the cells are formatted. The column header has comments.
	'==========================================================================
	
	response.write("Tutorial 07<br>")
	response.write("----------<br>")




	'Create an instance of the object that generates Excel files
	set xls = Server.CreateObject("EasyXLS.ExcelDocument")
	
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
	
	'Lock the first tab
	xls.easy_getSheetAt(0).setSheetProtected(true)
	
	'Get the table of the first worksheet
	Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
	
	'Create the style for the header
	set xlsStyleHeader = Server.CreateObject("EasyXLS.ExcelStyle")
	xlsStyleHeader.setFont("Verdana")
	xlsStyleHeader.setFontSize(8)
	xlsStyleHeader.setItalic(True)
	xlsStyleHeader.setBold(True)
	xlsStyleHeader.setForeground(CLng(COLOR_YELLOW))
	xlsStyleHeader.setBackground(CLng(COLOR_BLACK))
	xlsStyleHeader.setBorderColors CLng(COLOR_GRAY), CLng(COLOR_GRAY), CLng(COLOR_GRAY), CLng(COLOR_GRAY)
	xlsStyleHeader.setBorderStyles BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM, BORDER_BORDER_MEDIUM
	xlsStyleHeader.setHorizontalAlignment(ALIGNMENT_ALIGNMENT_CENTER)
	xlsStyleHeader.setVerticalAlignment(ALIGNMENT_ALIGNMENT_BOTTOM)
	xlsStyleHeader.setWrap(True)
	xlsStyleHeader.setDataType(DATATYPE_STRING)
	
	'Add the cells for header
	for column = 0 to 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (column + 1))
		xlsFirstTable.easy_getCell(0,column).setStyle(xlsStyleHeader)

		'Add comment
		xlsFirstTable.easy_getCell(0, column).setComment_2("This is column no " & (column + 1))
	next
	xlsFirstTable.easy_getRowAt(0).setHeight(30)
	
	'Create a style for cells
	Set xlsStyleData = Server.CreateObject("EasyXLS.ExcelStyle")
	xlsStyleData.setHorizontalAlignment(ALIGNMENT_ALIGNMENT_LEFT)
	xlsStyleData.setForeground(CLng(COLOR_DARKGRAY))
	xlsStyleData.setWrap(False)
	xlsStyleData.setLocked(True)
	xlsStyleData.setDataType(DATATYPE_STRING)
	
	'Add the cells for data
	for row = 0 to 99
		for column = 0 to 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setStyle(xlsStyleData)
		next
	next

	'Set column widths
	xlsFirstTable.setColumnWidth_2 0, 70
	xlsFirstTable.setColumnWidth_2 1, 100
	xlsFirstTable.setColumnWidth_2 2, 70
	xlsFirstTable.setColumnWidth_2 3, 100
	xlsFirstTable.setColumnWidth_2 4, 70

	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial07.xls<br>")
	xls.easy_WriteXLSFile ("C:\Samples\Tutorial07.xls")
	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Dispose memory
	xls.Dispose
%>
