'----------------------------------------------------------------
' Tutorial 27
'
' This tutorial shows how to encrypt and set the password required for opening a document
'-----------------------------------------------------------------

Imports EasyXLS

Module Tutorial27

    Sub Main()


        Console.WriteLine("Tutorial 27" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files, having 2 sheets
        Dim xls As New ExcelDocument(2)

        'Set the sheet names
        xls.easy_getSheetAt(0).setSheetName("First tab")
        xls.easy_getSheetAt(1).setSheetName("Second tab")

        'Set the password required for opening the document
        xls.easy_getOptions().setPasswordToOpen("password")

        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial27.xls.")
        xls.easy_WriteXLSFile("C:\\Samples\\Tutorial27.xls")

        'Confirm generation
        Dim sError As String = xls.easy_getError()
        If (sError.Equals("")) Then
            Console.Write(vbCrLf & "File successfully created. Press Enter to Exit...")
        Else
            Console.Write(vbCrLf & "Error encountered: " & sError & vbCrLf & "Press Enter to Exit...")
        End If

        'Dispose memory
        xls.Dispose()

        Console.ReadLine()
    End Sub

End Module
