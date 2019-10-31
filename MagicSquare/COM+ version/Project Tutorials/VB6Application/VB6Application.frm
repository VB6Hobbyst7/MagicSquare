VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   3645
   ClientTop       =   2730
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleMode       =   0  'User
   ScaleWidth      =   22325.08
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkNBHours 
      Caption         =   "NB Hours"
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   6960
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkOTHours 
      Caption         =   "OT Hours"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   6720
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkRegular 
      Caption         =   "Regular"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   6480
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkEstimated 
      Caption         =   "Estimated"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   6240
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkTask 
      Caption         =   "Task"
      Enabled         =   0   'False
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   6000
      Value           =   2  'Grayed
      Width           =   1335
   End
   Begin VB.CommandButton btnExportToExcel 
      Caption         =   "Export To Excel"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   7440
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3000
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   5292
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   " Generate chart with the following columns:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   5640
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "* sample data set source; totals are computed using formulas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   8535
   End
   Begin VB.Label Label3 
      Caption         =   "* sample hyperlink"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "* sample image"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label hlkEasyXLS 
      Caption         =   "http://www.easyxls.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "VB6Application.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Image imgEasyXLSlogo 
      Height          =   645
      Left            =   120
      Picture         =   "VB6Application.frx":030A
      Tag             =   "EasyXLSlogo.jpg"
      Top             =   120
      Width           =   2070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim varGetRows() As Variant
Dim rsDataGridSource As ADOR.Recordset

Private Sub btnExportToExcel_Click()
    
    ' Creating an instance of the object that generates excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")

    ' Adding a sheet to the Excel Document object
    Set xlsWorksheet = CreateObject("EasyXLS.ExcelWorksheet")
    xlsWorksheet.setSheetName ("TimeSheetReport")
    xls.easy_addWorksheet (xlsWorksheet)

    ' Adding the image
    xlsWorksheet.easy_addImage_5 imgEasyXLSlogo.Tag, "A1"
    
    ' Adding the hyperlink
    xlsWorksheet.easy_addHyperlink_2 HyperlinkType.HYPERLINKTYPE_URL, hlkEasyXLS.Caption, "A5"


    'Create an instance of the object used to format the cells
    Dim xlsAutoFormat
    Set xlsAutoFormat = CreateObject("EasyXLS.ExcelAutoFormat")
    xlsAutoFormat.InitAs (Styles.AUTOFORMAT_EASYXLS1)


    'Create the list used to store the values
    Dim lstRows
    Set lstRows = CreateObject("EasyXLS.Util.List")
    
    'Add the header row to the list
    Dim lstHeaderRow
    Set lstHeaderRow = CreateObject("EasyXLS.Util.List")
    lstHeaderRow.addElement ("Project")
    lstHeaderRow.addElement ("Resource")
    lstHeaderRow.addElement ("Role")
    lstHeaderRow.addElement ("Task")
    lstHeaderRow.addElement ("Estimated")
    lstHeaderRow.addElement ("Regular")
    lstHeaderRow.addElement ("OT Hours")
    lstHeaderRow.addElement ("NB Hours")
    lstHeaderRow.addElement ("Approval Status")
    lstRows.addElement (lstHeaderRow)
    
    ' Add the cells to the list
    rsDataGridSource.MoveFirst
    varGetRows = rsDataGridSource.GetRows()
    For nRowIndex = 0 To rsDataGridSource.Fields().Count
        Set RowList = CreateObject("EasyXLS.Util.List")
        For nColumnIndex = 0 To 8
            RowList.addElement ("" & varGetRows(nColumnIndex, nRowIndex))
        Next
        lstRows.addElement (RowList)
    Next
    
    ' Adding the content of the grid
    xlsWorksheet.easy_insertList_4 lstRows, 6, 0, xlsAutoFormat

    ' Creating the footer
    Dim nFooterRowIndex
    nFooterRowIndex = 6 + rsDataGridSource.Fields().Count + 1
    Dim xlsTable
    Set xlsTable = xlsWorksheet.easy_getExcelTable()
    xlsTable.easy_getCell(nFooterRowIndex, 0).setValue ("Totals:")
    xlsTable.easy_getCell(nFooterRowIndex, 4).setValue ("=SUM(E8:E" & nFooterRowIndex & ")")
    xlsTable.easy_getCell(nFooterRowIndex, 4).setDataType (DataType.DATATYPE_AUTOMATIC)
    xlsTable.easy_getCell(nFooterRowIndex, 5).setValue ("=SUM(F8:F" & nFooterRowIndex & ")")
    xlsTable.easy_getCell(nFooterRowIndex, 5).setDataType (DataType.DATATYPE_AUTOMATIC)
    xlsTable.easy_getCell(nFooterRowIndex, 6).setValue ("=SUM(G8:G" & nFooterRowIndex & ")")
    xlsTable.easy_getCell(nFooterRowIndex, 6).setDataType (DataType.DATATYPE_AUTOMATIC)
    xlsTable.easy_getCell(nFooterRowIndex, 7).setValue ("=SUM(H8:H" & nFooterRowIndex & ")")
    xlsTable.easy_getCell(nFooterRowIndex, 7).setDataType (DataType.DATATYPE_AUTOMATIC)
       
    Dim xlsChart
    Set xlsChart = CreateObject("EasyXLS.Charts.ExcelChart")
    xlsChart.setLeftUpperCorner_2 ("A20")
    xlsChart.setSize 600, 300
    If (chkEstimated.Value = 1) Then xlsChart.easy_addSeries_2 "=TimeSheetReport!$E$7", "=TimeSheetReport!$E$8:$E$16"
    If (chkRegular.Value = 1) Then xlsChart.easy_addSeries_2 "=TimeSheetReport!$F$7", "=TimeSheetReport!$F$8:$F$16"
    If (chkOTHours.Value = 1) Then xlsChart.easy_addSeries_2 "=TimeSheetReport!$G$7", "=TimeSheetReport!$G$8:$G$16"
    If (chkNBHours.Value = 1) Then xlsChart.easy_addSeries_2 "=TimeSheetReport!$H$7", "=TimeSheetReport!$H$8:$H$16"

    If (chkEstimated.Value = 1 Or chkRegular.Value = 1 Or chkOTHours.Value = 1 Or chkNBHours.Value = 1) Then
        xlsChart.easy_setCategoryXAxisLabels ("=TimeSheetReport!$D$8:$D$16")
    Else
        xlsChart.easy_addSeries_2 "=TimeSheetReport!$D$7", "=TimeSheetReport!$D$8:$D$16"
    End If
        
    xlsWorksheet.easy_addChart (xlsChart)
    

    'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\VB6Application.xls."
    xls.easy_WriteXLSFile ("c:\Samples\VB6Application.xls")

    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    
    'Dispose memory
    xls.Dispose
End Sub



Private Sub Form_Load()
    Styles.Initialize
    HyperlinkType.Initialize
    DataType.Initialize
        
    ' Create a new disconnected recordset object
    Set rsDataGridSource = New ADOR.Recordset
    
    rsDataGridSource.Fields.Append "Project", adVarChar, 50
    rsDataGridSource.Fields.Append "Resource", adVarChar, 50
    rsDataGridSource.Fields.Append "Role", adVarChar, 50
    rsDataGridSource.Fields.Append "Task", adVarChar, 50
    rsDataGridSource.Fields.Append "Estimated", adBigInt
    rsDataGridSource.Fields.Append "Regular", adBigInt
    rsDataGridSource.Fields.Append "OT Hours", adBigInt
    rsDataGridSource.Fields.Append "NB Hours", adBigInt
    rsDataGridSource.Fields.Append "Approval Status", adVarChar, 50
       
    
    rsDataGridSource.CursorType = adOpenDynamic
    rsDataGridSource.Open
    Dim lstFields
    lstFields = Array("Project", "Resource", "Role", "Task", "Estimated", "Regular", "OT Hours", "NB Hours", "Approval Status")
    
    Dim lstValues
    lstValues = Array("EasyXLS", "Jim Bean", "Programmer", "Build Charts", 800, 240, 40, 0, "To be Approved")
    rsDataGridSource.AddNew lstFields, lstValues
    lstValues = Array("EasyXLS", "Jack White", "Programmer", "Build Worksheets", 1000, 160, 0, 0, "To be Approved")
    rsDataGridSource.AddNew lstFields, lstValues
    lstValues = Array("EasyXLS", "Christina Brown", "Programmer", "Build Hyperlinks", 750, 256, 2, 0, "To be Approved")
    rsDataGridSource.AddNew lstFields, lstValues
    lstValues = Array("EasyXLS", "Walt Whitman", "Programmer", "Create Tutorials", 600, 114, 10, 0, "To be Approved")
    rsDataGridSource.AddNew lstFields, lstValues
    lstValues = Array("EasyXLS", "Adam Wilson", "Tester", "Test Charts", 120, 8, 0, 0, "To be Approved")
    rsDataGridSource.AddNew lstFields, lstValues
    lstValues = Array("EasyXLS", "Will Crane", "Tester", "Test Hyperlinks", 100, 10, 2, 0, "To be Approved")
    rsDataGridSource.AddNew lstFields, lstValues
    lstValues = Array("EasyXLS", "George Brown", "Artist", "Design", 300, 150, 2, 0, "To be Approved")
    rsDataGridSource.AddNew lstFields, lstValues
    lstValues = Array("MS Excel", "Christian Wurm", "Programmer", "Database Design", 120, 35, 3, 0, "To be Approved")
    rsDataGridSource.AddNew lstFields, lstValues
    lstValues = Array("MS Excel", "Adrian Fisher", "Tester", "Speed", 240, 48, 0, 8, "To be Approved")
    rsDataGridSource.AddNew lstFields, lstValues
    
   

    rsDataGridSource.MoveFirst
    varGetRows = rsDataGridSource.GetRows()
    
    rsDataGridSource.AddNew
    rsDataGridSource.Fields(0) = "Totals:"
    
    ' Computing the totals
    Dim nTotal
    nTotal = 0
    For nColumnIndex = 4 To 7
        nTotal = 0
        For nRowIndex = 0 To rsDataGridSource.Fields().Count - 1
            nTotal = nTotal + varGetRows(nColumnIndex, nRowIndex)
        Next
        rsDataGridSource.Fields(nColumnIndex) = nTotal
    Next

    
    ' Populate the datagrid
    Set DataGrid1.DataSource = rsDataGridSource
End Sub


Private Sub hlkEasyXLS_Click()
    Call ShellExecute(0&, vbNullString, hlkEasyXLS.Caption, vbNullString, vbNullString, vbNormalFocus)
End Sub



