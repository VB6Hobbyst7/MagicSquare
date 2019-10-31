VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   100
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    '==========================================================================
    ' Tutorial 20
    '
    '  This tutorial shows how to create a Microsoft Excel file
    ' that has AutoFilter.
    '==========================================================================
    
    DataType.Initialize
      
        
    Me.Label1.Caption = "Tutorial 20" & vbCrLf & "-----------------" & vbCrLf
    
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Create the worksheets
    xls.easy_addWorksheet_2 ("Sheet1")
       
    'Get the table of the first worksheet
    Set xlsTab = xls.easy_getSheet("Sheet1")
    Set xlsTable = xlsTab.easy_getExcelTable()
    
    'Add the cells for header
    For Column = 0 To 4
        xlsTable.easy_getCell(0, Column).setValue ("Column " & (Column + 1))
        xlsTable.easy_getCell(0, Column).setDataType (DataType.DATATYPE_STRING)
    Next
    
    'Add the cells for data
    For row = 0 To 99
        For Column = 0 To 4
            xlsTable.easy_getCell(row + 1, Column).setValue ("Data " & (row + 1) & ", " & (Column + 1))
            xlsTable.easy_getCell(row + 1, Column).setDataType (DataType.DATATYPE_STRING)
        Next
    Next
        
     'Add AutoFilter
    Set xlsFilter = xlsTab.easy_getFilter()
    xlsFilter.setAutoFilter_2 ("A1:E1")

    'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial20.xls"
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial20.xls")
    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    'Dispose memory
    xls.Dispose
End Sub

