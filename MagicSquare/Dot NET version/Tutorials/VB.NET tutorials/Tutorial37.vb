'----------------------------------------------------------------
' Tutorial 37
'
' This tutorial shows how to load a XLSX file (we use the file
' generated in Tutorial 28), modify some data and save it to
' another file (Tutorial37.xlsx).
'-----------------------------------------------------------------

Imports EasyXLS
Imports System.IO

Module Tutorial37

    Sub Main()


        Console.WriteLine("Tutorial 37" & vbCrLf & "----------" & vbCrLf)

        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument

        'Read the file
        Console.WriteLine("Reading file C:\Samples\Tutorial28.xlsx." & vbCrLf)
        If (xls.easy_LoadXLSXFile("C:\Samples\Tutorial28.xlsx")) Then

            'Get the table of the second worksheet
            Dim xlsSecondTab As ExcelWorksheet = xls.easy_getSheet("Second tab")
            Dim xlsSecondTable = xlsSecondTab.easy_getExcelTable

            'Write some data
            xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial37")

            For column As Integer = 0 To 4
                xlsSecondTable.easy_getCell(1, column).setValue("Data " & (column + 1))
            Next


            'Generate the file
            Console.WriteLine(vbCrLf & "Writing file C:\Samples\Tutorial37.xlsx.")
            xls.easy_WriteXLSXFile("C:\Samples\Tutorial37.xlsx")

            'Confirm generation
            Dim sError As String = xls.easy_getError()
            If (sError.Equals("")) Then
                Console.Write(vbCrLf & "File successfully created.")
            Else
                Console.Write(vbCrLf & "Error encountered: " & sError)
            End If
        Else
            Console.WriteLine(vbCrLf & "Error reading file C:\Samples\Tutorial28.xlsx " & vbCrLf & xls.easy_getError())
        End If

        'Dispose memory
        xls.Dispose()

        Console.WriteLine(vbCrLf & "Press Enter to Exit...")
        Console.ReadLine()
    End Sub

End Module
