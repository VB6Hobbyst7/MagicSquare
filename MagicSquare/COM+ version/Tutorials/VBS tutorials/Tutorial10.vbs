    '==========================================================================
    ' Tutorial 10
    '
    ' This tutorial shows how to merge a cell range.
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 10" & vbcrlf & "-----------" & vbcrlf)
    
   
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Add the worksheet
	xls.easy_addWorksheet_2("Sheet1")
	
	'Get the table of the first sheet
	Set xlsTable = xls.easy_getSheet("Sheet1").easy_getExcelTable()
	
	'Merging cells
	xlsTable.easy_mergeCells_2("A1:C3")
	    
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial10.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial10.xls")
    
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