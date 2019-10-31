    '==========================================================================
    'Tutorial 13
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The second one contains a named
    ' range. The first 10 rows of the first 2 columns contain
    ' validators.
    '==========================================================================

	'Constants declaration
	Dim VALIDATE_LIST, VALIDATE_WHOLE_NUMBER, OPERATOR_BETWEEN, OPERATOR_EQUAL_TO
    VALIDATE_LIST = 3
    VALIDATE_WHOLE_NUMBER = 1
    OPERATOR_BETWEEN = 0
    OPERATOR_EQUAL_TO = 2
    
    WScript.StdOut.WriteLine("Tutorial 13" & vbcrlf & "----------" & vbcrlf)
    
    
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
	xlsSecondTab.easy_addName_2 "Range", "=Second tab!$A$1:$A$4"

	'Add a validator for the first 10 rows of the first column
	set xlsFirstTab = xls.easy_getSheetAt(0)
	xlsFirstTab.easy_addDataValidator_3 "A1:A10", VALIDATE_LIST, OPERATOR_EQUAL_TO, "=Range", ""

	'Add a validator for the first 10 rows of the second column
	xlsFirstTab.easy_addDataValidator_3 "B1:B10", VALIDATE_WHOLE_NUMBER, OPERATOR_BETWEEN, "=4", "=100"
	
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial13.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial13.xls")
    
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
    