    '==========================================================================
    ' Tutorial 34
    '
    ' This tutorial shows how to read values from the active sheet
    ' of an excel file (the file generated in tutorial 9).
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 34" & vbcrlf & "-----------" & vbcrlf)
    
  


	'Create an instance of the object that generates Excel files
	set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Read the file
    WScript.StdOut.WriteLine("Reading file: C:\Samples\Tutorial09.xls")
    WScript.StdOut.WriteLine()
    Set rows = xls.easy_ReadXLSActiveSheet_AsList("C:\Samples\Tutorial09.xls")
    
    
    if xls.easy_getError() = "" then
		'Display the values
		For rowIndex = 0 To rows.Size() - 1
			Set row = rows.elementAt(rowIndex)
			For cellIndex = 0 To row.Size - 1
				WScript.StdOut.WriteLine("At row " & (rowIndex + 1) & ", column " & (cellIndex + 1) & " the value is '" & row.elementAt(cellIndex) & "'")
			Next
		Next
    else
		WScript.StdOut.Write(vbcrlf & "Error reading file C:\Samples\Tutorial09.xls " & xls.easy_getError())
    end if
    
    'Dispose memory
	xls.Dispose

	Wscript.StdOut.WriteLine()
    Wscript.StdOut.Write("Press Enter to exit ...")
    Wscript.StdIn.ReadLine
