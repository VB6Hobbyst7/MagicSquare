'----------------------------------------------------------------
' Tutorial 34
'
' This tutorial shows how to read values from the active sheet
' of an excel file (the file generated in Tutorial 09).
'-----------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants
Imports System.IO

Module Tutorial34

    Sub Main()


        Console.WriteLine("Tutorial 34" & vbCrLf & "----------" & vbCrLf)

        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument

        'Read the file
        Console.WriteLine("Reading file C:\Samples\Tutorial09.xls." & vbCrLf)
        Try
            Dim ds As DataSet = xls.easy_ReadXLSActiveSheet_AsDataSet("C:\Samples\Tutorial09.xls")

            'Display the values
            Dim dt As DataTable = ds.Tables(0)
            For row As Integer = 0 To dt.Rows.Count - 1
                For column As Integer = 0 To dt.Columns.Count - 1
                    Console.WriteLine("At row " & (row + 1) & ", column " & (column + 1) & _
                     " the value is '" & dt.Rows(row).ItemArray(column) & "'")
                Next
            Next
        Catch ex As Exception
            Console.WriteLine(vbCrLf & "Error reading file C:\Samples\Tutorial09.xls " & vbCrLf & xls.easy_getError())
        End Try

        'Dispose memory
        xls.Dispose()

        Console.Write(vbCrLf & "Press Enter to Exit...")
        Console.ReadLine()

    End Sub

End Module
