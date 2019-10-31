'----------------------------------------------------------------
' Tutorial 19
'
' This tutorial shows how to create a Microsoft Excel file
' that has two worksheets. The first one is full with data 
' and the first cell of the second row contains Rich Text Format.
'-----------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants


Module Tutorial19

    Sub Main()


        Console.WriteLine("Tutorial 19" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files, having 2 sheets
        Dim xls As New ExcelDocument(2)

        'Set the sheet names
        xls.easy_getSheetAt(0).setSheetName("First tab")
        xls.easy_getSheetAt(1).setSheetName("Second tab")

        'Get the table of the first worksheet
        Dim xlsFirstTab As ExcelWorksheet = xls.easy_getSheetAt(0)
        Dim xlsFirstTable As ExcelTable = xlsFirstTab.easy_getExcelTable()

        'Create the string used to set the RTF in cell
        Dim sFormattedValue As String
        sFormattedValue = "This is <b>bold</b>."
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
        xlsFirstTable.easy_getCell(1, 0).setHTMLValue(sFormattedValue)
        xlsFirstTable.easy_getCell(1, 0).setDataType(DataType.STRING)
        xlsFirstTable.easy_getCell(1, 0).setWrap(True)
        xlsFirstTable.easy_getRowAt(1).setHeight(250)
        xlsFirstTable.easy_getColumnAt(0).setWidth(250)




        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial19.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial19.xls")

        'Confirm generation
        Dim sError As String = xls.easy_getError()
        If (sError.Equals("")) Then
            Console.Write(vbCrLf & "File successfully created. Press Enter to Exit...")
        Else
            Console.Write(vbCrLf & "Error encountered: " & sError & vbCrLf & "Press Enter to Exit...")
        End If
        Console.ReadLine()
    End Sub

End Module
