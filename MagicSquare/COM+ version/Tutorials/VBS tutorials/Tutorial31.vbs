    '==========================================================================
    'Tutorial 31
    '
    ' This tutorial shows how to export an HTML file.
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 31" & vbcrlf & "----------" & vbcrlf)
    
	'Constants declaration
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"
    Dim AUTOFORMAT_EASYXLS1
    AUTOFORMAT_EASYXLS1 = 43

    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheet
	xls.easy_addWorksheet_2("First tab")
    
	'Get the table of the first worksheet
	Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
	
	'Add the cells for header
	For Column = 0 To 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (Column + 1))
		xlsFirstTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	Next

	'Add the cells for data
	For row = 0 To 99
		For column = 0 To 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		Next
	Next

	'Create an instance of the object used to format the cells
	Dim xlsAutoFormat 
	set xlsAutoFormat = CreateObject("EasyXLS.ExcelAutoFormat")
	xlsAutoFormat.InitAs(AUTOFORMAT_EASYXLS1)

	'Apply the predefined format to the cells.
	xlsFirstTable.easy_setRangeAutoFormat_2 "A1:E101", xlsAutoFormat


    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial31.html")
    xls.easy_WriteHTMLFile_3 "C:\Samples\Tutorial31.html","First tab"
    
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
    
