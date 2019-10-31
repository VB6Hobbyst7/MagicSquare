'----------------------------------------------------------------
' Tutorial 10
'
' This tutorial shows how to merge a cell range.
'-----------------------------------------------------------------

Imports EasyXLS

Module Tutorial10

    Sub Main()


        Console.WriteLine("Tutorial 10" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument(1)

        'Get the table of the first worksheet
        Dim xlsFirstTab As ExcelWorksheet = xls.easy_getSheet("Sheet1")
        Dim xlsTable = xlsFirstTab.easy_getExcelTable()

        'Merging cells
        xlsTable.easy_mergeCells("A1:C3")


        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial10.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial10.xls")

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
