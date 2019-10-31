'----------------------------------------------------------------
' Tutorial 41
'
' This tutorial shows how to load an XML file (we use the file
' generated in Tutorial 32), modify some data and save it to
' another file (Tutorial41.xls).
'-----------------------------------------------------------------

Imports EasyXLS
Imports System.IO

Module Tutorial41

    Sub Main()


        Console.WriteLine("Tutorial 41" & vbCrLf & "----------" & vbCrLf)

        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument

        'Read the file
        Console.WriteLine("Reading file C:\Samples\Tutorial32.xml." & vbCrLf)
        If (xls.easy_LoadXMLSpreadsheetFile("C:\Samples\Tutorial32.xml")) Then

            'Get the table of the second worksheet
            Dim xlsSecondTab As ExcelWorksheet = xls.easy_getSheetAt(1)
            Dim xlsTable = xlsSecondTab.easy_getExcelTable

            xlsTable.easy_getCell("A1").setValue("Data added by Tutorial41")

            For column As Integer = 0 To 4
                xlsTable.easy_getCell(1, column).setValue("Data " & (column + 1))
            Next


            'Generate the file
            Console.WriteLine(vbCrLf & "Writing file C:\Samples\Tutorial41.xls.")
            xls.easy_WriteXLSFile("C:\Samples\Tutorial41.xls")

            'Confirm generation
            Dim sError As String = xls.easy_getError()
            If (sError.Equals("")) Then
                Console.Write(vbCrLf & "File successfully created.")
            Else
                Console.Write(vbCrLf & "Error encountered: " & sError)
            End If
        Else
            Console.WriteLine(vbCrLf & "Error reading file C:\Samples\Tutorial32.xml " & vbCrLf & xls.easy_getError())
        End If

        'Dispose memory
        xls.Dispose()

        Console.WriteLine(vbCrLf & "Press Enter to Exit...")
        Console.ReadLine()
    End Sub

End Module
