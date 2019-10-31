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
    'Tutorial 11
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has a formula.
    '==========================================================================
    
    Me.Label1.Caption = "Tutorial 11" & vbCrLf & "---------------" & vbCrLf
    
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Add the worksheet
    xls.easy_addWorksheet_2 ("Formula")
    
    'Get the table, populate the sheet and set a formula
    Set xlsTable = xls.easy_getSheet("Formula").easy_getExcelTable()
    xlsTable.easy_getCell_2("A1").setValue ("1")
    xlsTable.easy_getCell_2("A2").setValue ("2")
    xlsTable.easy_getCell_2("A3").setValue ("3")
    xlsTable.easy_getCell_2("A4").setValue ("4")
    xlsTable.easy_getCell_2("A6").setValue ("=SUM(A1:A4)")

   
    'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial11.xls"
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial11.xls")
    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    'Dispose memory
    xls.Dispose
End Sub


