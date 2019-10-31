    '==========================================================================
    ' Tutorial 27
    '
    ' This tutorial shows how to encrypt and set the password required for opening a document
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 27" & vbcrlf & "----------" & vbcrlf)
    

    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
	
	'Set the password required for opening the document
	xls.easy_getOptions().setPasswordToOpen("password")
    
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial27.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial27.xls")
    
    'Confirm generation
    if sError = "" then
		WScript.StdOut.Write(vbcrlf & "File successfully created. Press Enter to exit...")
    else
		WScript.StdOut.Write(vbcrlf & "Error: " & xls.easy_getError())
    end if
    	
	'Dispose memory
	xls.Dispose
	
	WScript.StdIn.ReadLine()
    
