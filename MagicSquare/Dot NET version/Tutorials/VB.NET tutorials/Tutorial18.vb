'----------------------------------------------------------------
' Tutorial 18
'
' This tutorial shows how to create a Microsoft Excel file
' that has two worksheets. The first one is full with data 
' and the panes are frozen.
'-----------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants


Module Tutorial18

    Sub Main()


        Console.WriteLine("Tutorial 18" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files, having 2 sheets
        Dim xls As New ExcelDocument(2)

        'Set the sheet names
        xls.easy_getSheetAt(0).setSheetName("First tab")
        xls.easy_getSheetAt(1).setSheetName("Second tab")

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

        'Set column widths
        xlsFirstTable.setColumnWidth(0, 70)
        xlsFirstTable.setColumnWidth(1, 100)
        xlsFirstTable.setColumnWidth(2, 70)
        xlsFirstTable.setColumnWidth(3, 100)
        xlsFirstTable.setColumnWidth(4, 70)


        'Freeze panes
        xlsFirstTable.easy_freezePanes(1, 0, 75, 0)



        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial18.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial18.xls")

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
