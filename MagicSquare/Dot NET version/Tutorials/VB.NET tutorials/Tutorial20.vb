'----------------------------------------------------------------
' Tutorial 20
'
' This tutorial shows how to create a Microsoft Excel file 
' that has AutoFilter.

'-----------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants


Module Tutorial20

    Sub Main()


        Console.WriteLine("Tutorial 20" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files, having 1 sheet
        Dim xls As New ExcelDocument(1)

        'Set the sheet names
        xls.easy_getSheetAt(0).setSheetName("Sheet1")

        'Get the table of the first worksheet
        Dim xlsTab As ExcelWorksheet = xls.easy_getSheet("Sheet1")
        Dim xlsTable As ExcelTable = xlsTab.easy_getExcelTable()

        'Add the cells for header
        For column As Integer = 0 To 4
            xlsTable.easy_getCell(0, column).setValue("Column " & (column + 1))
            xlsTable.easy_getCell(0, column).setDataType(DataType.STRING)
        Next

        'Add the cells for data
        For row As Integer = 0 To 99
            For column As Integer = 0 To 4
                xlsTable.easy_getCell(row + 1, column).setValue("Data " & (row + 1) & ", " & (column + 1))
                xlsTable.easy_getCell(row + 1, column).setDataType(DataType.STRING)
            Next
        Next

        'Add AutoFilter
        Dim xlsFilter As ExcelFilter = xlsTab.easy_getFilter()
        xlsFilter.setAutoFilter("A1:E1")

        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial20.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial20.xls")

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
