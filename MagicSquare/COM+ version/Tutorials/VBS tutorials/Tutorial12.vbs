    '==========================================================================
    'Tutorial 12
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The second one contains a named
    ' range.
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 12" & vbcrlf & "----------" & vbcrlf)
    
   
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
    
	'Get the table of the second worksheet and populate the sheet
	set xlsSecondTab = xls.easy_getSheetAt(1)
	set xlsSecondTable = xlsSecondTab.easy_getExcelTable()
	xlsSecondTable.easy_getCell_2("A1").setValue("Range data 1")
	xlsSecondTable.easy_getCell_2("A2").setValue("Range data 2")
	xlsSecondTable.easy_getCell_2("A3").setValue("Range data 3")
	xlsSecondTable.easy_getCell_2("A4").setValue("Range data 4")

	'Create a named range
	xlsSecondTab.easy_addName_2 "Range", "='Second tab'!$A$1:$A$4"


    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial12.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial12.xls")
    
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
    