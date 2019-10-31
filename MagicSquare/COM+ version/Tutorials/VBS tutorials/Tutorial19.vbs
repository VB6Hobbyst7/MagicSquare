    '==========================================================================
    ' Tutorial 19
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The first one is full with data
    ' and the first cell of the second row contains Rich Text Format.
    '==========================================================================
    
    'Constants declaration
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"

    
    WScript.StdOut.WriteLine("Tutorial 19" & vbcrlf & "-----------" & vbcrlf)
    
   
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
    
	'Get the table of the first worksheet
	Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
	
	'Create the string used to set the RTF in cell
    Dim sFormattedValue 
    sFormattedValue = sFormattedValue & "This is <b>bold</b>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <i>italic</i>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <u>underline</u>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <underline double>double underline</underline double>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <font color=red>red</font>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <font color=rgb(255,0,0)>red</font> too."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <font face=""Arial Black"">Arial Black</font>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <font size=15pt>size 15</font>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <s>strikethrough</s>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <sup>superscript</sup>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <sub>subscript</sub>."
    sFormattedValue = sFormattedValue & Chr(10) & "<b>This</b> <i>is</i> <font color=red face=""Arial Black"" size=15pt><underline double>formatted</underline double></font> <s>text</s>."


    'Set the formatted value
    xlsFirstTable.easy_getCell(1, 0).setHTMLValue (sFormattedValue)
    xlsFirstTable.easy_getCell(1, 0).setDataType (DATATYPE_STRING)
    xlsFirstTable.easy_getCell(1, 0).setWrap (True)
    xlsFirstTable.easy_getRowAt(1).setHeight (250)
    xlsFirstTable.easy_getColumnAt(0).setWidth (250)


    
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial19.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial19.xls")
    
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
    