    '==========================================================================
    'Tutorial 11
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has a formula.
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 11" & vbcrlf & "----------" & vbcrlf)
    



	'Create an instance of the object that generates Excel files
	Set xls = CreateObject("EasyXLS.ExcelDocument")
	
	'Add one worksheet
	xls.easy_addWorksheet_2("Formula")
    
    'Get the table, populate the sheet and set a formula
    Set xlsTable = xls.easy_getSheet("Formula").easy_getExcelTable()
    xlsTable.easy_getCell_2("A1").setValue ("1")
    xlsTable.easy_getCell_2("A2").setValue ("2")
    xlsTable.easy_getCell_2("A3").setValue ("3")
    xlsTable.easy_getCell_2("A4").setValue ("4")
    xlsTable.easy_getCell_2("A6").setValue ("=SUM(A1:A4)")

    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial11.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial11.xls")
    
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
    