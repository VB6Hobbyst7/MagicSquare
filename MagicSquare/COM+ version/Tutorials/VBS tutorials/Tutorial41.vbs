    '==========================================================================
    ' Tutorial 41
    '
    ' This tutorial shows how to load an XML file (we use the file
    ' generated in tutorial 31), modify some data and save it to
    ' another file (Tutorial41.xls).
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 41" & vbcrlf & "-----------" & vbcrlf)
    

	'Create an instance of the object that generates Excel files
	set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Read the file
    WScript.StdOut.WriteLine("Reading file: C:\Samples\Tutorial32.xml" & vbcrlf)
    If (xls.easy_LoadXMLSpreadsheetFile_2("C:\Samples\Tutorial32.xml")) Then
    		
		'Get the table of the second worksheet and write some data
		Set xlsTable = xls.easy_getSheetAt(1).easy_getExcelTable()

		xlsTable.easy_getCell_2("A1").setValue("Data added by Tutorial41")

		for column=0 to 4
			xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
        
        'Generate the file
        Wscript.StdOut.WriteLine(vbcrlf & "Writing file: C:\Samples\Tutorial41.xls")
        xls.easy_WriteXLSFile ("C:\Samples\Tutorial41.xls")
    
		'Confirm generation
		dim sError
		sError = xls.easy_getError()
		if sError = "" then
		WScript.StdOut.WriteLine(vbcrlf & "File successfully created.")
		else
			WScript.StdOut.WriteLine(vbcrlf & "Error: " & sError)
		end if    
    Else
        Wscript.StdOut.WriteLine("Error reading file C:\Samples\Tutorial32.xml" & vbcrlf & xls.easy_getError())
    End If
    
	'Dispose memory
	xls.Dispose
    
    Wscript.StdOut.Write(vbcrlf & "Press Enter to exit ...")
    Wscript.StdIn.ReadLine()
