    '==========================================================================
    ' Tutorial 17
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The first one is full with data
    ' and contains groups.
    '==========================================================================
    
    'Constants declaration
    Dim AUTOFORMAT_EASYXLS1
    AUTOFORMAT_EASYXLS1 = 43
    Dim AUTOFORMAT_EASYXLS2
    AUTOFORMAT_EASYXLS2 = 45
    Dim DATAGROUP_GROUP_BY_ROWS
    DATAGROUP_GROUP_BY_ROWS = 0
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"

    
    WScript.StdOut.WriteLine("Tutorial 17" & vbcrlf & "-----------" & vbcrlf)
    
   
     'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
    
	'Get the table of the first worksheet
	Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
	
	'Add the cells for header
	For Column = 0 To 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (Column + 1))
		xlsFirstTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	Next
	xlsFirstTable.easy_getRowAt(0).setHeight(30)

	'Add the cells for data
	For row = 0 To 24
		For column = 0 To 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		Next
	Next

	'Set column widths
	xlsFirstTable.setColumnWidth_2 0, 70
	xlsFirstTable.setColumnWidth_2 1, 100
	xlsFirstTable.setColumnWidth_2 2, 70
	xlsFirstTable.setColumnWidth_2 3, 100
	xlsFirstTable.setColumnWidth_2 4, 70

        
    'Create the first group
    Set xlsFirstDataGroup = CreateObject("EasyXLS.ExcelDataGroup")
    xlsFirstDataGroup.setRange_2 ("A1:E26")
    xlsFirstDataGroup.setGroupType (DATAGROUP_GROUP_BY_ROWS)
    xlsFirstDataGroup.setCollapsed (False)
    
    'Create an instance of the object used to format the cells of the first group
    Dim xlsAutoFormat
    Set xlsAutoFormat = CreateObject("EasyXLS.ExcelAutoFormat")
    xlsAutoFormat.InitAs (AUTOFORMAT_EASYXLS1)
    xlsFirstDataGroup.setAutoFormat (xlsAutoFormat)
    xls.easy_getSheetAt(0).easy_addDataGroup (xlsFirstDataGroup)

    'Create the second group
    Set xlsSecondDataGroup = CreateObject("EasyXLS.ExcelDataGroup")
    xlsSecondDataGroup.setRange_2 ("A2:E10")
    xlsSecondDataGroup.setGroupType (DATAGROUP_GROUP_BY_ROWS)
    xlsSecondDataGroup.setCollapsed (False)
    
    'Create an instance of the object used to format the cells of the second group
    Dim xlsAutoFormat2
    Set xlsAutoFormat2 = CreateObject("EasyXLS.ExcelAutoFormat")
    xlsAutoFormat2.InitAs (AUTOFORMAT_EASYXLS2)
    xlsSecondDataGroup.setAutoFormat (xlsAutoFormat2)
    xls.easy_getSheetAt(0).easy_addDataGroup (xlsSecondDataGroup)

    
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial17.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial17.xls")
    
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
    