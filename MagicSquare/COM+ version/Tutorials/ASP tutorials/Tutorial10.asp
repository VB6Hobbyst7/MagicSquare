<%@ Language=VBScript %>

<%
'==========================================================================
' Tutorial 10
'
' This tutorial shows how to merge a cell range.
'==========================================================================

response.Write("Tutorial 10<br>")
response.Write("----------<br>")




'Create an instance of the object that generates Excel files
Set xls = Server.CreateObject("EasyXLS.ExcelDocument")

'Add the worksheet
xls.easy_addWorksheet_2("Sheet1")

'Get the table of the first sheet
Set xlsTable = xls.easy_getSheet("Sheet1").easy_getExcelTable()

'Merging cells
xlsTable.easy_mergeCells_2("A1:C3")

'Generate the file
response.Write("Writing file: C:\Samples\Tutorial10.xls<br>")
xls.easy_WriteXLSFile ("C:\Samples\Tutorial10.xls")

'Confirm generation
If xls.easy_getError() = "" Then
    response.Write("File successfully created.")
Else
    response.Write("Error encountered: " + xls.easy_getError())
End If

'Dispose memory
xls.Dispose
%>
