'----------------------------------------------------------------
' Tutorial 11
'
' This tutorial shows how to create a Microsoft Excel file
' that has a formula.
'-----------------------------------------------------------------

Imports EasyXLS

Module Tutorial11

    Sub Main()


        Console.WriteLine("Tutorial 11" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument

        'Add one worksheet
        xls.easy_addWorksheet("Formula")

        'Get the table, populate the sheet and set a formula
        Dim xlsFirstTab As ExcelWorksheet = xls.easy_getSheet("Formula")
        Dim xlsTable = xlsFirstTab.easy_getExcelTable()
        xlsTable.easy_getCell("A1").setValue("1")
        xlsTable.easy_getCell("A2").setValue("2")
        xlsTable.easy_getCell("A3").setValue("3")
        xlsTable.easy_getCell("A4").setValue("4")
        xlsTable.easy_getCell("A6").setValue("=SUM(A1:A4)")


        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial11.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial11.xls")

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
