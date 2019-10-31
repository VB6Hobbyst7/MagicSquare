'----------------------------------------------------------------
' Tutorial 12
'
' This tutorial shows how to create a Microsoft Excel file
' that has two worksheets. The second one contains a named
' range.
'-----------------------------------------------------------------

Imports EasyXLS

Module Tutorial12

    Sub Main()


        Console.WriteLine("Tutorial 12" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files, having 2 sheets
        Dim xls As New ExcelDocument(2)

        'Set the sheet names
        xls.easy_getSheetAt(0).setSheetName("First tab")
        xls.easy_getSheetAt(1).setSheetName("Second tab")

        'Get the table of the second worksheet and populate the sheet
        Dim xlsSecondTab As ExcelWorksheet = xls.easy_getSheetAt(1)
        Dim xlsSecondTable = xlsSecondTab.easy_getExcelTable()
        xlsSecondTable.easy_getCell("A1").setValue("Range data 1")
        xlsSecondTable.easy_getCell("A2").setValue("Range data 2")
        xlsSecondTable.easy_getCell("A3").setValue("Range data 3")
        xlsSecondTable.easy_getCell("A4").setValue("Range data 4")

        'Create a named range
        xlsSecondTab.easy_addName("Range", "='Second tab'!$A$1:$A$4")


        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial12.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial12.xls")

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
