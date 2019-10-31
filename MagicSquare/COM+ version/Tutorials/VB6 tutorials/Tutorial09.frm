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
    'Tutorial 09
    '
    ' This tutorial shows how to create a Microsoft Excel file
    ' that has two worksheets. The first one is full with data
    ' and the cells are formatted. The column header has comments.
    ' The first worksheet has header & footer. The print options are
    ' set for the first worksheet.
    '==========================================================================
    
    Alignment.Initialize
    Border.Initialize
    DataType.Initialize
    Color.Initialize
    Footer.Initialize
    Header.Initialize
    PageSetup.Initialize
        
    Me.Label1.Caption = "Tutorial 09" & vbCrLf & "---------------" & vbCrLf
    
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
    'Create the worksheets
    xls.easy_addWorksheet_2 ("First tab")
    xls.easy_addWorksheet_2 ("Second tab")
    
    'Lock the first tab
    xls.easy_getSheetAt(0).setSheetProtected (True)

    'Get the table of the first worksheet
    Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
    
    'Create the style for the header
    Set xlsStyleHeader = CreateObject("EasyXLS.ExcelStyle")
    xlsStyleHeader.setFont ("Verdana")
    xlsStyleHeader.setFontSize (8)
    xlsStyleHeader.setItalic (True)
    xlsStyleHeader.setBold (True)
    xlsStyleHeader.setForeground (CLng(Color.COLOR_YELLOW))
    xlsStyleHeader.setBackground (CLng(Color.COLOR_BLACK))
    xlsStyleHeader.setBorderColors CLng(Color.COLOR_GRAY), CLng(Color.COLOR_GRAY), CLng(Color.COLOR_GRAY), CLng(Color.COLOR_GRAY)
    xlsStyleHeader.setBorderStyles Border.BORDER_BORDER_MEDIUM, Border.BORDER_BORDER_MEDIUM, Border.BORDER_BORDER_MEDIUM, Border.BORDER_BORDER_MEDIUM
    xlsStyleHeader.setHorizontalAlignment (Alignment.ALIGNMENT_ALIGNMENT_CENTER)
    xlsStyleHeader.setVerticalAlignment (Alignment.ALIGNMENT_ALIGNMENT_BOTTOM)
    xlsStyleHeader.setWrap (True)
    xlsStyleHeader.setDataType (DataType.DATATYPE_STRING)

    'Add the cells for header
    For Column = 0 To 4
        xlsFirstTable.easy_getCell(0, Column).setValue ("Column " & (Column + 1))
        xlsFirstTable.easy_getCell(0, Column).setStyle (xlsStyleHeader)
                    
        'Add comment
        xlsFirstTable.easy_getCell(0, Column).setComment_2 ("This is column no " & (Column + 1))
    Next
    xlsFirstTable.easy_getRowAt(0).setHeight (30)
    
    'Create a style for cells
    Set xlsStyleData = CreateObject("EasyXLS.ExcelStyle")
    xlsStyleData.setHorizontalAlignment (Alignment.ALIGNMENT_ALIGNMENT_LEFT)
    xlsStyleData.setForeground (CLng(Color.COLOR_DARKGRAY))
    xlsStyleData.setWrap (False)
    xlsStyleData.setLocked (True)
    xlsStyleData.setDataType (DataType.DATATYPE_STRING)
    
    'Add the cells for data
    For row = 0 To 99
        For Column = 0 To 4
            xlsFirstTable.easy_getCell(row + 1, Column).setValue ("Data " & (row + 1) & ", " & (Column + 1))
            xlsFirstTable.easy_getCell(row + 1, Column).setStyle (xlsStyleData)
        Next
    Next

    'Set column widths
    xlsFirstTable.setColumnWidth_2 0, 70
    xlsFirstTable.setColumnWidth_2 1, 100
    xlsFirstTable.setColumnWidth_2 2, 70
    xlsFirstTable.setColumnWidth_2 3, 100
    xlsFirstTable.setColumnWidth_2 4, 70
    
    'Add headers for the first worksheet
    Set xlsFirstTab = xls.easy_getSheetAt(0)
    xlsFirstTab.easy_getHeaderAt_2(Header.HEADER_POSITION_CENTER).InsertSingleUnderline
    xlsFirstTab.easy_getHeaderAt_2(Header.HEADER_POSITION_CENTER).InsertFile
    xlsFirstTab.easy_getHeaderAt_2(Header.HEADER_POSITION_CENTER).InsertValue (" - How to create header and footer")

    xlsFirstTab.easy_getHeaderAt_2(Header.HEADER_POSITION_RIGHT).InsertDate
    xlsFirstTab.easy_getHeaderAt_2(Header.HEADER_POSITION_RIGHT).InsertValue (" ")
    xlsFirstTab.easy_getHeaderAt_2(Header.HEADER_POSITION_RIGHT).InsertTime

    'Add footer for the first worksheet
    xlsFirstTab.easy_getFooterAt_2(Footer.FOOTER_POSITION_CENTER).InsertPage
    xlsFirstTab.easy_getFooterAt_2(Footer.FOOTER_POSITION_CENTER).InsertValue (" of ")
    xlsFirstTab.easy_getFooterAt_2(Footer.FOOTER_POSITION_CENTER).InsertPages

    'Set Page Setup options
    Set xlsPageSetup = xlsFirstTab.easy_getPageSetup()
    xlsPageSetup.easy_setPrintArea_3 ("A1:E101")
    xlsPageSetup.easy_setRowsToRepeatAtTop_3 ("$1:$1")
    xlsPageSetup.setCenterHorizontally (True)
    xlsPageSetup.setOrientation (PageSetup.PAGESETUP_ORIENTATION_PORTRAIT)
    xlsPageSetup.setPageOrder (PageSetup.PAGESETUP_PAGE_ORDER_DOWN_THEN_OVER)
    xlsPageSetup.setPaperSize (PageSetup.PAGESETUP_PAPER_SIZE_A4)
    xlsPageSetup.setPrintComments (PageSetup.PAGESETUP_COMMENTS_AT_END_OF_SHEET)
    xlsPageSetup.setPrintGridlines (True)
    xlsFirstTable.easy_insertPageBreakAtRow (21)
    xlsFirstTable.easy_insertPageBreakAtRow (41)
    xlsFirstTable.easy_insertPageBreakAtRow (61)
    xlsFirstTable.easy_insertPageBreakAtRow (81)
    xlsFirstTab.setPageBreakPreview (True)

   
    'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial09.xls"
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial09.xls")
    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    'Dispose memory
    xls.Dispose
End Sub
