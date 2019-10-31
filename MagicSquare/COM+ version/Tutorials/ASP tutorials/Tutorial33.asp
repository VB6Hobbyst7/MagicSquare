<%@ Language=VBScript %>

<!-- #INCLUDE FILE="FileProperty.inc" -->
<%
'==========================================================================
' Tutorial 33
'
' This tutorial shows how to set the properties of the document.
'==========================================================================

response.Write("Tutorial 33<br>")
response.Write("----------<br>")




'Create an instance of the object that generates Excel files
Set xls = Server.CreateObject("EasyXLS.ExcelDocument")

'Add the worksheet
xls.easy_addWorksheet_2("Sheet1")

'Set the 'Subject' property
xls.getSummaryInformation().setSubject("This is the subject")

'Set the 'Manager' property
xls.getDocumentSummaryInformation().setManager("This is the manager")

'Set a custom property
xls.getDocumentSummaryInformation().setCustomProperty "PropertyName", VT_NUMBER, "4"

'Generate the file
response.Write("Writing file: C:\Samples\Tutorial33.xls<br>")
xls.easy_WriteXLSFile ("C:\Samples\Tutorial33.xls")

'Confirm generation
If xls.easy_getError() = "" Then
    response.Write("File successfully created.")
Else
    response.Write("Error encountered: " + xls.easy_getError())
End If

'Dispose memory
xls.Dispose
%>
