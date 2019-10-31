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
    ' Tutorial 15
    '
    ' This tutorial shows how to create a Hyperlink. There are 4
    ' types oh hyperlinks:
    '       1 - to an URL;
    '       2 - to a FILE;`
    '       3 - to a UNC;
    '       4 - to a CELL in the same file;
    '
    ' The link can be placed over multiple cels.
    '
    ' Every type of hyperlink accepts a tool tip description.
    '
    ' Every type of hyperlink accepts a text mark. A text mark is a
    ' link inside the file. Exemples:
    '       http://www.mysite.com/index.html#Chapter3
    '       c:\myfile.xls#Sheet2!D3
    '==========================================================================
    
    HyperlinkType.Initialize
    
    Me.Label1.Caption = "Tutorial 15" & vbCrLf & "-----------------" & vbCrLf
    
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
       'Create the worksheets
    xls.easy_addWorksheet_2 ("First tab")
    xls.easy_addWorksheet_2 ("Second tab")
    
    Set xlsTab1 = xls.easy_getSheetAt(0)
    Set xlsTab2 = xls.easy_getSheetAt(1)
    
    'Create the hyperlink to an URL
    xlsTab1.easy_addHyperlink_3 HyperlinkType.HYPERLINKTYPE_URL, "http://www.euoutsourcing.com", "Link to URL", "B2:E2"

    'Create the hyperlink to a FILE
    xlsTab1.easy_addHyperlink_3 HyperlinkType.HYPERLINKTYPE_FILE, "c:\tutorial27.xls", "Link to file", "B3"

    'Create the hyperlink to an UNC
    xlsTab1.easy_addHyperlink_3 HyperlinkType.HYPERLINKTYPE_UNC, "\\nicoar\samples\tutorial9.xls", "Link to UNC", "B4:D4"

    'Create the hyperlink to a CELL
    xlsTab1.easy_addHyperlink_3 HyperlinkType.HYPERLINKTYPE_CELL, "'Second tab'!D3", "Link to CELL", "B5"

    'Creating a name for the second sheet
    xlsTab2.easy_addName_2 "Name", "=Second tab!$A$1:$A$4"
    
    'Create the hyperlink to a name
    xlsTab1.easy_addHyperlink_3 HyperlinkType.HYPERLINKTYPE_CELL, "Name", "Link to a name", "B6"

    
     'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial15.xls"
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial15.xls")
    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    'Dispose memory
    xls.Dispose
End Sub
