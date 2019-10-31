    '================================================================================
    'Tutorial 29
    '
    ' This tutorial shows how to export an XLSB file that has multiple sheets in VBS. 
    ' The first sheet is filled with data.
    '================================================================================
    
    WScript.StdOut.WriteLine("Tutorial 29" & vbcrlf & "----------" & vbcrlf)
    
	'Constants declaration
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"

    
   
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheet
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")
    
	'Get the table of the first worksheet
	Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
	
	'Add the cells for header
	For Column = 0 To 4
		xlsFirstTable.easy_getCell(0,column).setValue("Column " & (Column + 1))
		xlsFirstTable.easy_getCell(0,column).setDataType(DATATYPE_STRING)
	Next

	'Add the cells for data
	For row = 0 To 99
		For column = 0 To 4
			xlsFirstTable.easy_getCell(row+1,column).setValue("Data " & (row + 1) & ", " & (column + 1))
			xlsFirstTable.easy_getCell(row+1,column).setDataType(DATATYPE_STRING)
		Next
	Next


    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial29.xlsb")
    xls.easy_WriteXLSBFile "C:\Samples\Tutorial29.xlsb"
    
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
    
