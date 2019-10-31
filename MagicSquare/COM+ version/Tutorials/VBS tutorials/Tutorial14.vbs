    '==========================================================================
    ' Tutorial 14
    '
    ' This tutorial shows how to create conditional formatting ranges.
    '==========================================================================
    
    'Constants declaration
    Dim Bisque, Red
    Bisque = &hffc4e4ff
    Red = &hff0000ff

	
    Dim DT_NUMERIC
    DT_NUMERIC = "numeric"
    
    Dim CONDITIONAL_FORMATTING_OPERATOR_BETWEEN, CONDITIONAL_FORMATTING_OPERATOR_EQUALTO, CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA
    CONDITIONAL_FORMATTING_OPERATOR_BETWEEN = 1
    CONDITIONAL_FORMATTING_OPERATOR_EQUALTO = 3
    CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA = 2

    
    WScript.StdOut.WriteLine("Tutorial 14" & vbcrlf & "-----------" & vbcrlf)
    
   
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheets
	xls.easy_addWorksheet_2("Sheet1")
    
	'Get the table of the second worksheet and populate the sheet
	set xlsTab = xls.easy_getSheet("Sheet1")
	set xlsTable = xlsTab.easy_getExcelTable()

    For i = 0 To 5
        For j = 0 To 3
			If ((i < 2) And (j < 2)) Then
                xlsTable.easy_getCell(i, j).setValue ("12")
            Else
                If ((j = 2) And (i < 2)) Then
                    xlsTable.easy_getCell(i, j).setValue ("1000")
                Else
                    xlsTable.easy_getCell(i, j).setValue ("9")
                End If
            End If
            xlsTable.easy_getCell(i, j).setDataType (DT_NUMERIC)
        Next
    Next

	'Set a conditional formatting
	xlsTab.easy_addConditionalFormatting_5 "A1:C3", CONDITIONAL_FORMATTING_OPERATOR_BETWEEN, "=9", "=11", true, true, Clng(RED)

	'Set a conditional formatting
	xlsTab.easy_addConditionalFormatting_9 "A6:C6", CONDITIONAL_FORMATTING_OPERATOR_BETWEEN, "=COS(PI())+2", "", Clng(BISQUE)
	xlsTab.easy_getConditionalFormattingAt_2("A6:C6").getConditionAt(0).setConditionType(CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA)
    
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial14.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial14.xls")
    
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
    