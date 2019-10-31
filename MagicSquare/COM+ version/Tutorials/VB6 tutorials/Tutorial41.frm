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
    ' Tutorial 41
    '
    ' This tutorial shows how to load an XML file (we use the file
    ' generated in Tutorial 32), modify some data and save it to
    ' another file (Tutorial41.xls).
    '==========================================================================
    
    Me.Label1.Caption = "Tutorial 41" & vbCrLf & "-----------------" & vbCrLf
    

    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Read the file
    Me.Label1.Caption = Me.Label1.Caption & "Reading file: C:\Samples\Tutorial32.xml" & vbCrLf
    If (xls.easy_LoadXMLSpreadsheetFile_2("C:\Samples\Tutorial32.xml")) Then
    
        'Get the table of the second worksheet and write some data
        Set xlsTable = xls.easy_getSheetAt(1).easy_getExcelTable()
        xlsTable.easy_getCell_2("A1").setValue ("Data added by Tutorial41")

        For Column = 0 To 4
            xlsTable.easy_getCell(1, Column).setValue ("Data " & (Column + 1))
        Next
        
        'Generate the file
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial41.xls"
        xls.easy_WriteXLSFile ("C:\Samples\Tutorial41.xls")
        
        'Confirm generation
        If xls.easy_getError() = "" Then
            Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
        Else
            Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
        End If
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error reading file C:\Samples\Tutorial32.xml" & vbCrLf & xls.easy_getError()
    End If
    
    'Dispose memory
    xls.Dispose
End Sub
