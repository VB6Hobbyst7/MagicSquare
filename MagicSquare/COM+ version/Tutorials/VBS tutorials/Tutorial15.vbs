    '==========================================================================
    ' Tutorial 15
    '
    ' This tutorial shows how to create a Hyperlink. There are 4
    ' types of hyperlinks:
    '       1 - to an URL;
    '       2 - to a FILE;
    '       3 - to a UNC;
    '       4 - to a CELL in the same file;
    '
    ' The link can be placed over multiple cells.
    '
    ' Every type of hyperlink accepts a tool tip description.
    '
    ' Every type of hyperlink accepts a text mark. A text mark is a
    ' link inside the file. Examples:
    '       http://www.mysite.com/index.html#Chapter3
    '       c:\myfile.xls#Sheet2!D3
    '==========================================================================
    
    'Constants declaration
    Dim HYPERLINKTYPE_URL, HYPERLINKTYPE_FILE, HYPERLINKTYPE_UNC, HYPERLINKTYPE_CELL
    HYPERLINKTYPE_URL = "url"
	HYPERLINKTYPE_FILE = "file"
	HYPERLINKTYPE_UNC = "unc"
	HYPERLINKTYPE_CELL = "cell"
    
    WScript.StdOut.WriteLine("Tutorial 15" & vbcrlf & "-----------" & vbcrlf)
    
   
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
	
	set xlsTab1 = xls.easy_getSheetAt(0)
	set xlsTab2 = xls.easy_getSheetAt(1)
    
	'Create the hyperlink to an URL
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_URL, "http://www.euoutsourcing.com", "Link to URL", "B2:E2"

	'Create the hyperlink to a FILE
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_FILE, "c:\myfile.xls", "Link to file", "B3"

	'Create the hyperlink to an UNC
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_UNC, "\\computerName\Folder\file.txt", "Link to UNC", "B4:D4"

	'Create the hyperlink to a CELL
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_CELL, "'Second tab'!D3", "Link to CELL", "B5"

	'Creating a name for the second sheet
	xlsTab2.easy_addName_2 "Name", "=Second tab!$A$1:$A$4"
	
	'Create the hyperlink to a name
	xlsTab1.easy_addHyperlink_3 HYPERLINKTYPE_CELL, "Name", "Link to a name", "B6"
	
	'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial15.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial15.xls")
    
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
    