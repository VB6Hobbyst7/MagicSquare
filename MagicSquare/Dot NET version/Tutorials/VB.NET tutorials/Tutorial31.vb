'----------------------------------------------------------------
' Tutorial 31
'
' This tutorial shows how to export an HTML file.
'-----------------------------------------------------------------*/

Imports EasyXLS
Imports EasyXLS.Constants

Module Tutorial31

    Sub Main()


        Console.WriteLine("Tutorial 31" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files, having 2 sheets
        Dim xls As New ExcelDocument(2)

        'Set the sheet name
        xls.easy_getSheetAt(0).setSheetName("First tab")

        'Get the table of the first worksheet
        Dim xlsFirstTab As ExcelWorksheet = xls.easy_getSheetAt(0)
        Dim xlsFirstTable = xlsFirstTab.easy_getExcelTable()

        'Add the cells for header
        For column As Integer = 0 To 4
            xlsFirstTable.easy_getCell(0, column).setValue("Column " & (column + 1))
            xlsFirstTable.easy_getCell(0, column).setDataType(DataType.STRING)
        Next
        xlsFirstTable.easy_getRowAt(0).setHeight(30)


        'Add the cells for data
        For row As Integer = 0 To 99
            For column As Integer = 0 To 4
                xlsFirstTable.easy_getCell(row + 1, column).setValue("Data " & (row + 1) & ", " & (column + 1))
                xlsFirstTable.easy_getCell(row + 1, column).setDataType(DataType.STRING)
            Next
        Next

        'Apply a predefined format to the cells.
        xlsFirstTable.easy_setRangeAutoFormat("A1:E101", New ExcelAutoFormat(Styles.AUTOFORMAT_EASYXLS1))

        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial31.html.")
        xls.easy_WriteHTMLFile("C:\Samples\Tutorial31.html", "First tab")

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
