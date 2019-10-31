'----------------------------------------------------------------
' Tutorial 36
'
' This tutorial shows how to load an excel file (we use the one
' generated in Tutorial 09), modify some data and save it to
' another file (Tutorial36.xls).
'-----------------------------------------------------------------

Imports EasyXLS
Imports System.IO

Module Tutorial36

    Sub Main()


        Console.WriteLine("Tutorial 36" & vbCrLf & "----------" & vbCrLf)

        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument

        'Read the file
        Console.WriteLine("Reading file C:\Samples\Tutorial09.xls." & vbCrLf)
        'Dim file As New FileStream("C:\Samples\Tutorial09.xls", FileMode.Open)
        If (xls.easy_LoadXLSFile("C:\Samples\Tutorial09.xls")) Then

            'Get the table of the second worksheet
            Dim xlsSecondTab As ExcelWorksheet = xls.easy_getSheet("Second tab")
            Dim xlsSecondTable = xlsSecondTab.easy_getExcelTable

            'Write some data
            xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial36")

            For column As Integer = 0 To 4
                xlsSecondTable.easy_getCell(1, column).setValue("Data " & (column + 1))
            Next


            'Generate the file
            Console.WriteLine(vbCrLf & "Writing file C:\Samples\Tutorial36.xls.")
            xls.easy_WriteXLSFile("C:\Samples\Tutorial36.xls")

            'Confirm generation
            Dim sError As String = xls.easy_getError()
            If (sError.Equals("")) Then
                Console.Write(vbCrLf & "File successfully created.")
            Else
                Console.Write(vbCrLf & "Error encountered: " & sError)
            End If
        Else
            Console.WriteLine(vbCrLf & "Error reading file C:\Samples\Tutorial09.xls " & vbCrLf & xls.easy_getError())
        End If

        'Dispose memory
        xls.Dispose()

        Console.WriteLine(vbCrLf & "Press Enter to Exit...")
        Console.ReadLine()
    End Sub

End Module
