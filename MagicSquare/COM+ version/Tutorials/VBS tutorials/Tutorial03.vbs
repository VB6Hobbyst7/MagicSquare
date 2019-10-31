    '==========================================================================
    ' Tutorial 03
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets.
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 03" & vbcrlf & "----------" & vbcrlf)
    

    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
    
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial03.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial03.xls")
    
    'Confirm generation
    if sError = "" then
		WScript.StdOut.Write(vbcrlf & "File successfully created. Press Enter to exit...")
    else
		WScript.StdOut.Write(vbcrlf & "Error: " & xls.easy_getError())
    end if
    	
	'Dispose memory
	xls.Dispose
	
	WScript.StdIn.ReadLine()
    