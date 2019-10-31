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
    ' Tutorial 13
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The second one contains a named
    ' range. The first 10 rows of the first 2 columns contain
    ' validators.
    '==========================================================================
    
    DataValidator.Initialize
    
    Me.Label1.Caption = "Tutorial 13" & vbCrLf & "-----------------" & vbCrLf
    
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Create the worksheets
    xls.easy_addWorksheet_2 ("First tab")
    xls.easy_addWorksheet_2 ("Second tab")
    
    'Get the table of the second worksheet and populate the sheet
    Set xlsSecondTab = xls.easy_getSheetAt(1)
    Set xlsSecondTable = xlsSecondTab.easy_getExcelTable()
    xlsSecondTable.easy_getCell_2("A1").setValue ("Range data 1")
    xlsSecondTable.easy_getCell_2("A2").setValue ("Range data 2")
    xlsSecondTable.easy_getCell_2("A3").setValue ("Range data 3")
    xlsSecondTable.easy_getCell_2("A4").setValue ("Range data 4")

    'Create a named range
    xlsSecondTab.easy_addName_2 "Range", "=Second tab!$A$1:$A$4"
   
   'Add a validator for the first 10 rows of the first column
    Set xlsFirstTab = xls.easy_getSheetAt(0)
    xlsFirstTab.easy_addDataValidator_3 "A1:A10", DataValidator.DATAVALIDATOR_VALIDATE_LIST, DataValidator.DATAVALIDATOR_OPERATOR_EQUAL_TO, "=Range", ""

    'Add a validator for the first 10 rows of the second column
    xlsFirstTab.easy_addDataValidator_3 "B1:B10", DataValidator.DATAVALIDATOR_VALIDATE_WHOLE_NUMBER, DataValidator.DATAVALIDATOR_OPERATOR_BETWEEN, "=4", "=100"


    'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial13.xls"
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial13.xls")
    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    'Dispose memory
    xls.Dispose
End Sub
