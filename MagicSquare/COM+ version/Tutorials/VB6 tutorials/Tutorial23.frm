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
    ' Tutorial 23
    '
    ' This tutorial shows how to modify different properties
    ' related to the chart.
    '==========================================================================
    
    Format.Initialize
    Color.Initialize
    Chart.Initialize
    LineStyleFormat.Initialize
    ShadowFormat.Initialize
    
    Me.Label1.Caption = "Tutorial 23" & vbCrLf & "-----------------" & vbCrLf


    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
   'Add one worksheet
    xls.easy_addWorksheet_2 ("SourceData")
    
    '----------------------------------------------------------------------
    'Insert values
    Set xlsTable1 = xls.easy_getSheet("SourceData").easy_getExcelTable()

    xlsTable1.easy_getCell(0, 0).setValue ("Show Date")
    xlsTable1.easy_getCell(0, 1).setValue ("Available Places")
    xlsTable1.easy_getCell(0, 2).setValue ("Available Tickets")
    xlsTable1.easy_getCell(0, 3).setValue ("Sold Tickets")

    xlsTable1.easy_getCell(1, 0).setValue ("03/13/2005 00:00:00")
    xlsTable1.easy_getCell(1, 0).setFormat (Format.FORMAT_FORMAT_DATE)
    xlsTable1.easy_getCell(2, 0).setValue ("03/14/2005 00:00:00")
    xlsTable1.easy_getCell(2, 0).setFormat (Format.FORMAT_FORMAT_DATE)
    xlsTable1.easy_getCell(3, 0).setValue ("03/15/2005 00:00:00")
    xlsTable1.easy_getCell(3, 0).setFormat (Format.FORMAT_FORMAT_DATE)
    xlsTable1.easy_getCell(4, 0).setValue ("03/16/2005 00:00:00")
    xlsTable1.easy_getCell(4, 0).setFormat (Format.FORMAT_FORMAT_DATE)
    
    xlsTable1.easy_getCell(1, 1).setValue ("10000")
    xlsTable1.easy_getCell(2, 1).setValue ("5000")
    xlsTable1.easy_getCell(3, 1).setValue ("8500")
    xlsTable1.easy_getCell(4, 1).setValue ("1000")

    xlsTable1.easy_getCell(1, 2).setValue ("8000")
    xlsTable1.easy_getCell(2, 2).setValue ("4000")
    xlsTable1.easy_getCell(3, 2).setValue ("6000")
    xlsTable1.easy_getCell(4, 2).setValue ("1000")

    xlsTable1.easy_getCell(1, 3).setValue ("920")
    xlsTable1.easy_getCell(2, 3).setValue ("1005")
    xlsTable1.easy_getCell(3, 3).setValue ("342")
    xlsTable1.easy_getCell(4, 3).setValue ("967")

    xlsTable1.easy_getColumnAt(0).setWidth (100)
    xlsTable1.easy_getColumnAt(1).setWidth (100)
    xlsTable1.easy_getColumnAt(2).setWidth (100)
    xlsTable1.easy_getColumnAt(3).setWidth (100)


    '--------------------------------------------------------------------------
   
    'Add the chart
    xls.easy_addChart_5 "Chart", "=SourceData!$A$1:$D$5", Chart.CHART_SERIES_IN_COLUMNS

    'Get the previously added chart
    Set xlsChartSheet = xls.easy_getSheetAt(1)
    Set xlsChart = xlsChartSheet.easy_getExcelChart()

    'Modifying chart type
    xlsChart.easy_setChartType (Chart.CHART_CHART_TYPE_CYLINDER_COLUMN)

    'Modifying chart area properties
    Set xlsChartArea = xlsChart.easy_getChartArea()
    xlsChartArea.getLineColorFormat().setLineColor (CLng(Color.COLOR_DARKGRAY))
    xlsChartArea.getLineStyleFormat().setDashType (LineStyleFormat.LINESTYLEFORMAT_DASH_TYPE_SOLID)
    xlsChartArea.getLineStyleFormat().setWidth (0.25)
    
    'Modifying chart plot area properties
    Set xlsPlotArea = xlsChart.easy_getPlotArea()
    xlsPlotArea.getLineColorFormat().setLineColor (CLng(Color.COLOR_DARKGRAY))
    xlsPlotArea.getLineStyleFormat().setDashType (LineStyleFormat.LINESTYLEFORMAT_DASH_TYPE_SOLID)
    xlsPlotArea.getLineStyleFormat().setWidth (0.25)

    'Modifying legend property
    Set xlsChatLegend = xlsChart.easy_getLegend()
    xlsChatLegend.getFillFormat().setBackground (CLng(Color.COLOR_LAVENDERBLUSH))
    xlsChatLegend.getFontFormat().setForeground (CLng(Color.COLOR_BLUE))
    xlsChatLegend.getFontFormat().setItalic (True)
    xlsChatLegend.setKeysArrangementDirection (Chart.CHART_KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL)
    xlsChatLegend.setPlacement (Chart.CHART_LEGEND_CORNER)
    xlsChatLegend.getShadowFormat().setShadow (ShadowFormat.SHADOWFORMAT_OFFSET_DIAGONAL_BOTTOM_RIGHT)

    'Modifying X axis properties
    Set xlsXAxis = xlsChart.easy_getCategoryXAxis()
    xlsXAxis.getLineColorFormat().setLineColor (CLng(Color.COLOR_STEELBLUE))
    xlsXAxis.getLineStyleFormat().setDashType (LineStyleFormat.LINESTYLEFORMAT_DASH_TYPE_DASH_DOT)
    xlsXAxis.getLineStyleFormat().setWidth (0.25)
    xlsXAxis.getFontFormat().setForeground (CLng(Color.COLOR_RED))

    'Modifying Y axis properties
    Set xlsYAxis = xlsChart.easy_getValueYAxis()
    xlsYAxis.getLineColorFormat().setLineColor (CLng(Color.COLOR_STEELBLUE))
    xlsYAxis.getLineStyleFormat().setDashType (LineStyleFormat.LINESTYLEFORMAT_DASH_TYPE_LONG_DASH)
    xlsYAxis.getLineStyleFormat().setWidth (0.25)
    xlsYAxis.getFontFormat().setForeground (CLng(Color.COLOR_BLUE))

    'Modifying series properties
    xlsChart.easy_getSeriesAt(0).getFillFormat().setBackground (CLng(Color.COLOR_ROYALBLUE))
    xlsChart.easy_getSeriesAt(1).getFillFormat().setBackground (CLng(Color.COLOR_YELLOW))
    xlsChart.easy_getSeriesAt(2).getFillFormat().setBackground (CLng(Color.COLOR_LIGHTGREEN))
    

    'Generate the file
    Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Writing file C:\Samples\Tutorial23.xls"
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial23.xls")
    
    'Confirm generation
    If xls.easy_getError() = "" Then
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "File successfully created."
    Else
        Me.Label1.Caption = Me.Label1.Caption & vbCrLf & "Error: " & xls.easy_getError()
    End If

    'Dispose memory
    xls.Dispose
End Sub
