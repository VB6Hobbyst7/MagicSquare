    '==========================================================================
    ' Tutorial 39
    '
    ' This tutorial shows how to load a CSV file (we use the file
    ' generated in tutorial 29), modify some data and save it to
    ' another file (Tutorial39.xls).
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 39" & vbcrlf & "-----------" & vbcrlf)
    
  


	'Create an instance of the object that generates Excel files
	set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Read the file
    WScript.StdOut.WriteLine("Reading file: C:\Samples\Tutorial30.csv" & vbcrlf)
    If (xls.easy_LoadCSVFile("C:\Samples\Tutorial30.csv")) Then
    
		'Set the name of the first worksheet
		xls.easy_getSheetAt(0).setSheetName("First tab")
		

		'Add a new worksheet and write some data
		xls.easy_addWorksheet_2("Second tab")
		Set xlsTable = xls.easy_getSheetAt(1).easy_getExcelTable()

		xlsTable.easy_getCell_2("A1").setValue("Data added by Tutorial39")

		for column=0 to 4
			xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
        
        'Generate the file
        Wscript.StdOut.WriteLine(vbcrlf & "Writing file: C:\Samples\Tutorial39.xls")
        xls.easy_WriteXLSFile ("C:\Samples\Tutorial39.xls")
    
		'Confirm generation
		dim sError
		sError = xls.easy_getError()
		if sError = "" then
		WScript.StdOut.WriteLine(vbcrlf & "File successfully created.")
		else
			WScript.StdOut.WriteLine(vbcrlf & "Error: " & sError)
		end if    
    Else
        Wscript.StdOut.WriteLine("Error reading file C:\Samples\Tutorial30.csv" & vbcrlf & xls.easy_getError())
    End If
    
	'Dispose memory
	xls.Dispose
    
    Wscript.StdOut.Write(vbcrlf & "Press Enter to exit ...")
    Wscript.StdIn.ReadLine()
