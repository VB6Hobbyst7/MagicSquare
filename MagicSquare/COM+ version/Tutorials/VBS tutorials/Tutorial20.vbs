    '==========================================================================
    ' Tutorial 20
    '
    ' This tutorial shows how to create a Microsoft Excel file 
    ' that has AutoFilter.

    '==========================================================================
    
    'Constants declaration
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"

    
    WScript.StdOut.WriteLine("Tutorial 20" & vbcrlf & "-----------" & vbcrlf)
    
   
     'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
     'Create the worksheets
	xls.easy_addWorksheet_2("Sheet1")
	    
	'Get the table of the first worksheet
      Set xlsTab = xls.easy_getSheet("Sheet1")
	Set xlsTable = xlsTab.easy_getExcelTable()
	
	'Add the cells for header
	For Column = 0 To 4
		xlsTable.easy_getCell(0,column).setValue("Column " & (Column + 1))
		xlsTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	Next
	
	'Add the cells for data
	For row = 0 To 99
		For column = 0 To 4
			xlsTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		Next
	Next

	'Add AutoFilter
      Set xlsFilter = xlsTab.easy_getFilter()
      xlsFilter.setAutoFilter_2("A1:E1")

    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial20.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial20.xls")
    
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
    