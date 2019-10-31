'----------------------------------------------------------------
' Tutorial 38
'
' This tutorial shows how to load a XLSB file (we use the file
' generated in Tutorial 29), modify some data and save it to
' another file (Tutorial38.xlsb).
'-----------------------------------------------------------------

Imports EasyXLS
Imports System.IO

Module Tutorial38

    Sub Main()


        Console.WriteLine("Tutorial 38" & vbCrLf & "----------" & vbCrLf)

        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument

        'Read the file
        Console.WriteLine("Reading file C:\Samples\Tutorial29.xlsb." & vbCrLf)
        If (xls.easy_LoadXLSBFile("C:\Samples\Tutorial29.xlsb")) Then

            'Get the table of the second worksheet
            Dim xlsSecondTab As ExcelWorksheet = xls.easy_getSheet("Second tab")
            Dim xlsSecondTable = xlsSecondTab.easy_getExcelTable

            'Write some data
            xlsSecondTable.easy_getCell("A1").setValue("Data added by Tutorial38")

            For column As Integer = 0 To 4
                xlsSecondTable.easy_getCell(1, column).setValue("Data " & (column + 1))
            Next


            'Generate the file
            Console.WriteLine(vbCrLf & "Writing file C:\Samples\Tutorial38.xlsb.")
            xls.easy_WriteXLSBFile("C:\Samples\Tutorial38.xlsb")

            'Confirm generation
            Dim sError As String = xls.easy_getError()
            If (sError.Equals("")) Then
                Console.Write(vbCrLf & "File successfully created.")
            Else
                Console.Write(vbCrLf & "Error encountered: " & sError)
            End If
        Else
            Console.WriteLine(vbCrLf & "Error reading file C:\Samples\Tutorial29.xlsb " & vbCrLf & xls.easy_getError())
        End If

        'Dispose memory
        xls.Dispose()

        Console.WriteLine(vbCrLf & "Press Enter to Exit...")
        Console.ReadLine()
    End Sub

End Module
