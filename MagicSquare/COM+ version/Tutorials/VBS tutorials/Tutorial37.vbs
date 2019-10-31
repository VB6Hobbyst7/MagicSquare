    '==========================================================================
    ' Tutorial 37
    '
    ' This tutorial shows how to load a XLSX file (we use the file
    ' generated in tutorial 27), modify some data and save it to
    ' another file (Tutorial37.xlsx).
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 37" & vbcrlf & "-----------" & vbcrlf)
    
  


	'Create an instance of the object that generates Excel files
	set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Read the file
    WScript.StdOut.WriteLine("Reading file: C:\Samples\Tutorial28.xlsx" & vbcrlf)
    If (xls.easy_LoadXLSXFile("C:\Samples\Tutorial28.xlsx")) Then
    
		'Get the table of the second worksheet
		Set xlsSecondTable = xls.easy_getSheet("Second tab").easy_getExcelTable()

		'Write some data
		xlsSecondTable.easy_getCell_2("A1").setValue("Data added by Tutorial37")

		for column=0 to 4
			xlsSecondTable.easy_getCell(1, column).setValue("Data " & (column + 1))
		next
        
        'Generate the file
        Wscript.StdOut.WriteLine(vbcrlf & "Writing file: C:\Samples\Tutorial37.xlsx")
        xls.easy_WriteXLSXFile ("C:\Samples\Tutorial37.xlsx")
    
		'Confirm generation
		dim sError
		sError = xls.easy_getError()
		if sError = "" then
		WScript.StdOut.WriteLine(vbcrlf & "File successfully created.")
		else
			WScript.StdOut.WriteLine(vbcrlf & "Error: " & sError)
		end if    
    Else
        Wscript.StdOut.WriteLine("Error reading file C:\Samples\Tutorial28.xlsx" & vbcrlf & xls.easy_getError())
    End If
    
	'Dispose memory
	xls.Dispose
    
    Wscript.StdOut.Write(vbcrlf & "Press Enter to exit ...")
    Wscript.StdIn.ReadLine()
