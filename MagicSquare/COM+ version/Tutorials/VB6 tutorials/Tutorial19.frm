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
    ' Tutorial 19
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The first one is full with data
    ' and the first cell of the second row contains Rich Text Format.
    '==========================================================================
    
    DataType.Initialize
          
        
    Me.Label1.Caption = "Tutorial 19" & vbCrLf & "-----------------" & vbCrLf
    
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Create the worksheets
    xls.easy_addWorksheet_2 ("First tab")
    xls.easy_addWorksheet_2 ("Second tab")
    
    'Get the table of the first worksheet
    Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
    
    'Create the string used to set the RTF in cell
    Dim sFormattedValue As String
    sFormattedValue = sFormattedValue & "This is <b>bold</b>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <i>italic</i>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <u>underline</u>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <underline double>double underline</underline double>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <font color=red>red</font>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <font color=rgb(255,0,0)>red</font> too."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <font face=""Arial Black"">Arial Black</font>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <font size=15pt>size 15</font>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <s>strikethrough</s>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <sup>superscript</sup>."
    sFormattedValue = sFormattedValue & Chr(10) & "This is <sub>subscript</sub>."
    sFormattedValue = sFormattedValue & Chr(10) & "<b>This</b> <i>is</i> <font color=red face=""Arial Black"" size=15pt><underline double>formatted</underline double></font> <s>text</s>."


    'Set the formatted value
    xlsFirstTable.easy_getCell(1, 0).setHTMLValue (sFormattedValue)
    xlsFirstTable.easy_getCell(1, 0).setDataType (DataType.DATATYPE_STRING)
    xlsFirstTable.easy_getCell(1, 0).setWrap (True)
    xlsFirstTable.easy_getRowAt(1).setHeight (250)
    xlsFirstTable.easy_getColumnAt(0).setWidth (250)


        


    'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial19.xls"
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial19.xls")
    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    'Dispose memory
    xls.Dispose
End Sub

