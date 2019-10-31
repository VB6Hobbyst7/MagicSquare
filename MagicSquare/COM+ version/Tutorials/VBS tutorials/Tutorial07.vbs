    '==========================================================================
    'Tutorial 07
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The first one is full with data
    ' and the cells are formatted. The column header has comments.
    '==========================================================================
    
    'Constants declaration
	Dim DATATYPE_STRING
    DATATYPE_STRING = "string"

    Dim ALIGNMENT_CENTER, ALIGNMENT_BOTTOM, ALIGNMENT_LEFT
    ALIGNMENT_CENTER = "center"
    ALIGNMENT_BOTTOM = "bottom"
    ALIGNMENT_LEFT = "left"
    
    Dim Black, Gray, Yellow, DarkGray, Blue
    Black = &hff000000
    Gray = &hff808080
    Yellow = &hff00ffff
    DarkGray = &hffa9a9a9
    Blue = &hffff0000
    
    Dim BORDER_MEDIUM
    BORDER_MEDIUM = 2
        
    WScript.StdOut.WriteLine("Tutorial 07" & vbcrlf & "----------" & vbcrlf)
    
   
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
	
	'Lock the first tab
	xls.easy_getSheetAt(0).setSheetProtected(true)
    
	'Get the table of the first worksheet
	Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
    
	'Create the style for the header
	set xlsStyleHeader = CreateObject("EasyXLS.ExcelStyle")
	xlsStyleHeader.setFont("Verdana")
	xlsStyleHeader.setFontSize(8)
	xlsStyleHeader.setItalic(True)
	xlsStyleHeader.setBold(True)
	xlsStyleHeader.setForeground(CLng(YELLOW))
	xlsStyleHeader.setBackground(CLng(BLACK))
	xlsStyleHeader.setBorderColors CLng(GRAY), CLng(GRAY), CLng(GRAY), CLng(GRAY)
	xlsStyleHeader.setBorderStyles BORDER_MEDIUM, BORDER_MEDIUM, BORDER_MEDIUM, BORDER_MEDIUM
	xlsStyleHeader.setHorizontalAlignment(ALIGNMENT_CENTER)
	xlsStyleHeader.setVerticalAlignment(ALIGNMENT_BOTTOM)
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
	Set xlsStyleData = CreateObject("EasyXLS.ExcelStyle")
	xlsStyleData.setHorizontalAlignment(ALIGNMENT_LEFT)
	xlsStyleData.setForeground(CLng(DARKGRAY))
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
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial07.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial07.xls")
    
    'Confirm generation
    dim sError
    sError = xls.easy_getError()
    if sError = "" then
		WScript.StdOut.Write(vbcrlf & "File successfully created. Press Enter to exit...")
    else
		WScript.StdOut.Write(vbcrlf & "Error: " & sError)
    end if
    WScript.StdIn.ReadLine()
    
    'Dispose memory
	xls.Dispose   