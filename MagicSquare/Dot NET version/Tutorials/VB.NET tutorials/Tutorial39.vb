'----------------------------------------------------------------
' Tutorial 39
'
' This tutorial shows how to load a CSV file (we use the file
' generated in Tutorial 30), modify some data and save it to
' another file (Tutorial39.xls).
'-----------------------------------------------------------------

Imports EasyXLS
Imports System.IO

Module Tutorial39

    Sub Main()


        Console.WriteLine("Tutorial 39" & vbCrLf & "----------" & vbCrLf)

        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument

        'Read the file
        Console.WriteLine("Reading file C:\Samples\Tutorial30.csv." & vbCrLf)
        If (xls.easy_LoadCSVFile("C:\Samples\Tutorial30.csv")) Then

            'Set the name of the first worksheet
            xls.easy_getSheetAt(0).setSheetName("First tab")

            'Add a new worksheet and write some data
            xls.easy_addWorksheet("Second tab")
            Dim xlsSecondTab As ExcelWorksheet = xls.easy_getSheetAt(1)
            Dim xlsTable = xlsSecondTab.easy_getExcelTable

            xlsTable.easy_getCell("A1").setValue("Data added by Tutorial39")

            For column As Integer = 0 To 4
                xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
            Next


            'Generate the file
            Console.WriteLine(vbCrLf & "Writing file C:\Samples\Tutorial39.xls.")
            xls.easy_WriteXLSFile("C:\Samples\Tutorial39.xls")

            'Confirm generation
            Dim sError As String = xls.easy_getError()
            If (sError.Equals("")) Then
                Console.Write(vbCrLf & "File successfully created.")
            Else
                Console.Write(vbCrLf & "Error encountered: " & sError)
            End If
        Else
            Console.WriteLine(vbCrLf & "Error reading file C:\Samples\Tutorial30.csv " & vbCrLf & xls.easy_getError())
        End If

        'Dispose memory
        xls.Dispose()

        Console.WriteLine(vbCrLf & "Press Enter to Exit...")
        Console.ReadLine()
    End Sub

End Module
