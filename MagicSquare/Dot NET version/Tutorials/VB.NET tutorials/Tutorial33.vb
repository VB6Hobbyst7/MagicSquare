'----------------------------------------------------------------
' Tutorial 33
'
' This tutorial shows how to set the properties of the document.
'-----------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants


Module Tutorial33

    Sub Main()


        Console.WriteLine("Tutorial 33" & vbCrLf & "----------" & vbCrLf)

        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument(1)

        'Set the 'Subject' property
        xls.getSummaryInformation().setSubject("This is the subject")

        'Set the 'Manager' property
        xls.getDocumentSummaryInformation().setManager("This is the manager")

        'Set a custom property
        xls.getDocumentSummaryInformation().setCustomProperty("PropertyName", FileProperty.VT_NUMBER, "4")

        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial33.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial33.xls")

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
