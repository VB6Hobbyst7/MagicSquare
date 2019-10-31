    '==========================================================================
    ' Tutorial 33
    '
    'This tutorial shows how to set the properties of the document.
    '==========================================================================

	WScript.StdOut.WriteLine("Tutorial 33" & vbcrlf & "-----------" & vbcrlf)

    Dim VT_NUMBER
    VT_NUMBER = 5
	
	'Create an instance of the object that generates Excel files
	set xls = CreateObject("EasyXLS.ExcelDocument")

	'Add the worksheet
	xls.easy_addWorksheet_2("Sheet1")

	'Set the 'Subject' property
	xls.getSummaryInformation().setSubject("This is the subject")
	
	'Set the 'Manager' property
	xls.getDocumentSummaryInformation().setManager("This is the manager")

	'Set a custom property
	xls.getDocumentSummaryInformation().setCustomProperty "PropertyName", VT_NUMBER, "4"


        'Generate the file
        Wscript.StdOut.WriteLine(vbcrlf & "Writing file: C:\Samples\Tutorial33.xls")
        xls.easy_WriteXLSFile ("C:\Samples\Tutorial33.xls")

'Confirm generation
		dim sError
		sError = xls.easy_getError()
		if sError = "" then
		WScript.StdOut.WriteLine(vbcrlf & "File successfully created.")
		else
			WScript.StdOut.WriteLine(vbcrlf & "Error: " & sError)
		end if   

'Dispose memory
	xls.Dispose
    
    Wscript.StdOut.Write(vbcrlf & "Press Enter to exit ...")
    Wscript.StdIn.ReadLine()
