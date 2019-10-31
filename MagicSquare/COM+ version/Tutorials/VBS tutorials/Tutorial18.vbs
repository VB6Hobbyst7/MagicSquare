    '==========================================================================
    ' Tutorial 18
    '
	' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The first one is full with data
    ' and the panes are frozen.
    '==========================================================================
    
    'Constants declaration
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"

    
    WScript.StdOut.WriteLine("Tutorial 18" & vbcrlf & "-----------" & vbcrlf)
    
   
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
	For row = 0 To 99
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


   'Freeze panes
    xlsFirstTable.easy_freezePanes_2 1, 0, 75, 0


    
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial18.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial18.xls")
    
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
    