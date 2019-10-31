'----------------------------------------------------------------
' Tutorial 16
'
' This tutorial shows how to create a Microsoft Excel file
' that has two worksheets. The first one has an image.
'-----------------------------------------------------------------

Imports EasyXLS

Module Tutorial16

    Sub Main()


        Console.WriteLine("Tutorial 16" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files, having 2 sheets
        Dim xls As New ExcelDocument(2)

        'Set the sheet names
        xls.easy_getSheetAt(0).setSheetName("First tab")
        xls.easy_getSheetAt(1).setSheetName("Second tab")

        'Create the image
        Dim xlsFirstTab As ExcelWorksheet = xls.easy_getSheetAt(0)
        xlsFirstTab.easy_addImage("C:\\Samples\\EasyXLSLogo.JPG", "A1")


        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial16.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial16.xls")

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
