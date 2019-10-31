    '==========================================================================
    'Tutorial 16
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The first one has an image.
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 16" & vbcrlf & "----------" & vbcrlf)
    

   
	'Create an instance of the object that generates Excel files
	Set xls = CreateObject("EasyXLS.ExcelDocument")
	
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
	
	'Create the image
	xls.easy_getSheetAt(0).easy_addImage_5 "C:\\Samples\\EasyXLSLogo.JPG", "A1"
		
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial16.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial16.xls")
    
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
