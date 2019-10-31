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
    ' Tutorial 17
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The first one is full with data
    ' and contains groups.
    '==========================================================================
    
    DataType.Initialize
    Styles.Initialize
    DataGroup.Initialize
     
        
    Me.Label1.Caption = "Tutorial 17" & vbCrLf & "-----------------" & vbCrLf
    
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Create the worksheets
    xls.easy_addWorksheet_2 ("First tab")
    xls.easy_addWorksheet_2 ("Second tab")
    
    'Get the table of the first worksheet
    Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
    
    'Add the cells for header
    For Column = 0 To 4
        xlsFirstTable.easy_getCell(0, Column).setValue ("Column " & (Column + 1))
        xlsFirstTable.easy_getCell(0, Column).setDataType (DataType.DATATYPE_STRING)
    Next
    xlsFirstTable.easy_getRowAt(0).setHeight (30)

    'Add the cells for data
    For row = 0 To 24
        For Column = 0 To 4
            xlsFirstTable.easy_getCell(row + 1, Column).setValue ("Data " & (row + 1) & ", " & (Column + 1))
            xlsFirstTable.easy_getCell(row + 1, Column).setDataType (DataType.DATATYPE_STRING)
        Next
    Next

    'Set column widths
    xlsFirstTable.setColumnWidth_2 0, 70
    xlsFirstTable.setColumnWidth_2 1, 100
    xlsFirstTable.setColumnWidth_2 2, 70
    xlsFirstTable.setColumnWidth_2 3, 100
    xlsFirstTable.setColumnWidth_2 4, 70
        
    'Create the first group
    Set xlsFirstDataGroup = CreateObject("EasyXLS.ExcelDataGroup")
    xlsFirstDataGroup.setRange_2 ("A1:E26")
    xlsFirstDataGroup.setGroupType (DataGroup.DATAGROUP_GROUP_BY_ROWS)
    xlsFirstDataGroup.setCollapsed (False)
    
    'Create an instance of the object used to format the cells of the first group
    Dim xlsAutoFormat
    Set xlsAutoFormat = CreateObject("EasyXLS.ExcelAutoFormat")
    xlsAutoFormat.InitAs (Styles.AUTOFORMAT_EASYXLS1)
    xlsFirstDataGroup.setAutoFormat (xlsAutoFormat)
    xls.easy_getSheetAt(0).easy_addDataGroup (xlsFirstDataGroup)

    'Create the second group
    Set xlsSecondDataGroup = CreateObject("EasyXLS.ExcelDataGroup")
    xlsSecondDataGroup.setRange_2 ("A2:E10")
    xlsSecondDataGroup.setGroupType (DataGroup.DATAGROUP_GROUP_BY_ROWS)
    xlsSecondDataGroup.setCollapsed (False)
    
    'Create an instance of the object used to format the cells of the second group
    Dim xlsAutoFormat2
    Set xlsAutoFormat2 = CreateObject("EasyXLS.ExcelAutoFormat")
    xlsAutoFormat2.InitAs (Styles.AUTOFORMAT_EASYXLS2)
    xlsSecondDataGroup.setAutoFormat (xlsAutoFormat2)
    xls.easy_getSheetAt(0).easy_addDataGroup (xlsSecondDataGroup)
        


    'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial17.xls"
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial17.xls")
    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    'Dispose memory
    xls.Dispose
End Sub
