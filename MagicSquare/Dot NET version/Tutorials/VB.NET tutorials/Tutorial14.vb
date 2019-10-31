'----------------------------------------------------------------
' Tutorial 14
'
' This tutorial shows how to create conditional formatting ranges.
'-----------------------------------------------------------------

Imports System.Drawing
Imports EasyXLS
Imports EasyXLS.Constants

Module Tutorial14

    Sub Main()


        Console.WriteLine("Tutorial 14" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument(1)

        'Get the table of the first worksheet
        Dim xlsTab As ExcelWorksheet = xls.easy_getSheet("Sheet1")
        Dim xlsTable = xlsTab.easy_getExcelTable()

        'Insert data
        For i As Integer = 0 To 5
            For j As Integer = 0 To 3
                If ((i < 2) And (j < 2)) Then
                    xlsTable.easy_getCell(i, j).setValue("12")
                Else
                    If ((j = 2) And (i < 2)) Then
                        xlsTable.easy_getCell(i, j).setValue("1000")
                    Else
                        xlsTable.easy_getCell(i, j).setValue("9")
                    End If

                    xlsTable.easy_getCell(i, j).setDataType(DataType.NUMERIC)
                End If
            Next
        Next

        'Set a conditional formatting
        xlsTab.easy_addConditionalFormatting("A1:C3", ConditionalFormatting.OPERATOR_BETWEEN, "=9", "=11", True, True, Color.Red)

        'Set a conditional formatting
        xlsTab.easy_addConditionalFormatting("A6:C6", ConditionalFormatting.OPERATOR_BETWEEN, "=COS(PI())+2", "", Color.Bisque)
        xlsTab.easy_getConditionalFormattingAt("A6:C6").getConditionAt(0).setConditionType(ConditionalFormatting.CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA)


        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial14.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial14.xls")

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
