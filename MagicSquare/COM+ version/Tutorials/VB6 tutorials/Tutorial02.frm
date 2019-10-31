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
    ' Tutorial 02
    '
    ' This tutorial shows how to generate an Excel document from a list of values.
    ' The cells are formatted using an user-defined format.
    '==========================================================================
    
    Alignment.Initialize
    Color.Initialize

    'Constants declaration
    Dim OddRowStripesStyleColor
    OddRowStripesStyleColor = &HFFF0F7EF


    Me.Label1.Caption = "Tutorial 02" & vbCrLf & "---------------" & vbCrLf
    
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
   'Connect to the database
    Dim objConn
    Set objConn = CreateObject("ADODB.Connection")
    objConn.ConnectionString = "Provider=SQLOLEDB;Server=(local);Database=northwind;User ID=sa;Password=;"
    objConn.Open


    
    Dim sQueryString
    sQueryString = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', P.ProductName AS 'Product Name', O.UnitPrice AS Price, ' ' + cast(O.Quantity AS varchar) AS Quantity , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE  O.ProductID = P.ProductID AND O.OrderID = ord.OrderID"
    
    'Create the record set object
    Dim objRS
    Set objRS = CreateObject("ADODB.Recordset")
    objRS.Open sQueryString, objConn
    
    'Create the list used to store the values
    Dim lstRows
    Set lstRows = CreateObject("EasyXLS.Util.List")
    
    'Add the header row to the list
    Dim lstHeaderRow
    Set lstHeaderRow = CreateObject("EasyXLS.Util.List")
    lstHeaderRow.addElement ("Order Date")
    lstHeaderRow.addElement ("Product Name")
    lstHeaderRow.addElement ("Price")
    lstHeaderRow.addElement ("Quantity")
    lstHeaderRow.addElement ("Value")
    lstRows.addElement (lstHeaderRow)
    
    'Add the values from the database to the list
    Do Until objRS.EOF = True
        Set RowList = CreateObject("EasyXLS.Util.List")
        RowList.addElement ("" & objRS("Order Date"))
        RowList.addElement ("" & objRS("Product Name"))
        RowList.addElement ("" & objRS("Price"))
        RowList.addElement ("" & objRS("Quantity"))
        RowList.addElement ("" & objRS("Value"))
        lstRows.addElement (RowList)
            
       'Move to the next record
       objRS.MoveNext
    Loop
    

    'Create an instance of the object used to format the cells
    Dim xlsAutoFormat
    Set xlsAutoFormat = CreateObject("EasyXLS.ExcelAutoFormat")
    'Set the style of the header
    Dim xlsHeaderStyle
    Set xlsHeaderStyle = CreateObject("EasyXLS.ExcelStyle")
    xlsHeaderStyle.setBackground (CLng(Color.COLOR_LIGHTGREEN))
    xlsHeaderStyle.setFontSize (12)
    xlsAutoFormat.setHeaderRowStyle (xlsHeaderStyle)

    ' Set the style of the cells
    Dim xlsEvenRowStripesStyle
    Set xlsEvenRowStripesStyle = CreateObject("EasyXLS.ExcelStyle")
    xlsEvenRowStripesStyle.setBackground (CLng(Color.COLOR_FLORALWHITE))
    xlsEvenRowStripesStyle.setFormat ("$0.00")
    xlsEvenRowStripesStyle.setHorizontalAlignment (Alignment.ALIGNMENT_ALIGNMENT_LEFT)
    xlsAutoFormat.setEvenRowStripesStyle (xlsEvenRowStripesStyle)
    Dim xlsOddRowStripesStyle
    Set xlsOddRowStripesStyle = CreateObject("EasyXLS.ExcelStyle")
    xlsOddRowStripesStyle.setBackground (OddRowStripesStyleColor)
    xlsOddRowStripesStyle.setFormat ("$0.00")
    xlsOddRowStripesStyle.setHorizontalAlignment (Alignment.ALIGNMENT_ALIGNMENT_LEFT)
    xlsAutoFormat.setOddRowStripesStyle (xlsOddRowStripesStyle)
    Dim xlsLeftColumnStyle
    Set xlsLeftColumnStyle = CreateObject("EasyXLS.ExcelStyle")
    xlsLeftColumnStyle.setBackground (CLng(Color.COLOR_FLORALWHITE))
    xlsLeftColumnStyle.setFormat ("mm/dd/yyyy")
    xlsLeftColumnStyle.setHorizontalAlignment (Alignment.ALIGNMENT_ALIGNMENT_LEFT)
    xlsAutoFormat.setLeftColumnStyle (xlsLeftColumnStyle)
    

    'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial02.xls."
    xls.easy_WriteXLSFile_FromList_2 "c:\Samples\Tutorial02.xls", lstRows, xlsAutoFormat, "Sheet1"

    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    'Close the Recordset object
    objRS.Close
    
    'Delete the Recordset Object
    Set objRS = Nothing
    
    
    'Close the Connection object
    objConn.Close
    
    'Delete the Connection Object
    Set objConn = Nothing
    
    'Dispose memory
    xls.Dispose
End Sub
