    '==========================================================================
    ' Tutorial 23
    '
    ' This tutorial shows how to modify different properties
    ' related to the chart.
    '==========================================================================
    
    'Constants declaration
    Dim FORMAT_DATE
    FORMAT_DATE = "MM/dd/yyyy"
    
    Dim Blue, DarkGray, LavenderBlush, SteelBlue, Red, RoyalBlue, Yellow, LightGreen
    Blue = &hffff0000
    DarkGray = &hffa9a9a9
    LavenderBlush = &hfff5f0ff
    SteelBlue = &hffb48246
    Red = &hff0000ff
    RoyalBlue = &hffe16941
    Yellow = &hff00ffff
    LightGreen = &hff90ee90
    
    Dim CHART_TYPE_CYLINDER_COLUMN
    Dim CHART_SERIES_IN_COLUMNS	
    Dim DASH_TYPE_SOLID, DASH_TYPE_DASH_DOT, DASH_TYPE_LONG_DASH
    Dim KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL
    Dim LEGEND_CORNER, OFFSET_DIAGONAL_BOTTOM_RIGHT
    CHART_TYPE_CYLINDER_COLUMN = 110
    CHART_SERIES_IN_COLUMNS = 1
    DASH_TYPE_SOLID = "solid"
    DASH_TYPE_DASH_DOT = "dashDot"
    DASH_TYPE_LONG_DASH = "lgDash"
    KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL = 0
    LEGEND_CORNER = 1
    OFFSET_DIAGONAL_BOTTOM_RIGHT = 1
    
    WScript.StdOut.WriteLine("Tutorial 23" & vbcrlf & "-----------" & vbcrlf)
  
    'Create an instance of the object that generates Excel files
    set xls = CreateObject("EasyXLS.ExcelDocument")
	
    'Add one worksheet
    xls.easy_addWorksheet_2("SourceData")
	
    '----------------------------------------------------------------------
    'Insert values
    Set xlsTable1 = xls.easy_getSheet("SourceData").easy_getExcelTable()

    xlsTable1.easy_getCell(0, 0).setValue ("Show Date")
    xlsTable1.easy_getCell(0, 1).setValue ("Available Places")
    xlsTable1.easy_getCell(0, 2).setValue ("Available Tickets")
    xlsTable1.easy_getCell(0, 3).setValue ("Sold Tickets")

    xlsTable1.easy_getCell(1, 0).setValue ("03/13/2005 00:00:00")
    xlsTable1.easy_getCell(1, 0).setFormat (FORMAT_DATE)
    xlsTable1.easy_getCell(2, 0).setValue ("03/14/2005 00:00:00")
    xlsTable1.easy_getCell(2, 0).setFormat (FORMAT_DATE)
    xlsTable1.easy_getCell(3, 0).setValue ("03/15/2005 00:00:00")
    xlsTable1.easy_getCell(3, 0).setFormat (FORMAT_DATE)
    xlsTable1.easy_getCell(4, 0).setValue ("03/16/2005 00:00:00")
    xlsTable1.easy_getCell(4, 0).setFormat (FORMAT_DATE)
    
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
    xls.easy_addChart_5 "Chart", "=SourceData!$A$1:$D$5", CHART_SERIES_IN_COLUMNS

    'Get the previously added chart    
	Set xlsChartSheet = xls.easy_getSheetAt(1)
	Set xlsChart = xlsChartSheet.easy_getExcelChart()

    'Modifying chart type
    xlsChart.easy_setChartType (CHART_TYPE_CYLINDER_COLUMN)

    'Modifying chart area properties
    Set xlsChartArea = xlsChart.easy_getChartArea()
    xlsChartArea.getLineColorFormat().setLineColor(CLng(DarkGray))
    xlsChartArea.getLineStyleFormat().setDashType (DASH_TYPE_SOLID)
    xlsChartArea.getLineStyleFormat(). setWidth(0.25)
    
    'Modifying chart plot area properties
    Set xlsPlotArea = xlsChart.easy_getPlotArea()
    xlsPlotArea.getLineColorFormat().setLineColor (CLng(DarkGray))
    xlsPlotArea.getLineStyleFormat().setDashType (DASH_TYPE_SOLID)
    xlsPlotArea.getLineStyleFormat().setWidth(0.25)

    'Modifying legend property
    Set xlsChartLegend = xlsChart.easy_getLegend()
    xlsChartLegend.getFillFormat().setBackground (CLng(LavenderBlush))
    xlsChartLegend.getFontFormat().setForeground (CLng(Blue))
    xlsChartLegend.getFontFormat().setItalic (True)
    xlsChartLegend.setKeysArrangementDirection (KEYS_ARRANGEMENT_DIRECTION_HORIZONTAL)
    xlsChartLegend.setPlacement (LEGEND_CORNER)
    xlsChartLegend.getShadowFormat().setShadow(OFFSET_DIAGONAL_BOTTOM_RIGHT)

    'Modifying X axis properties
    Set xlsXAxis = xlsChart.easy_getCategoryXAxis()
    xlsXAxis.getLineColorFormat().setLineColor (CLng(SteelBlue))
    xlsXAxis.getLineStyleFormat().setDashType (DASH_TYPE_DASH_DOT)
    xlsXAxis.getLineStyleFormat().setWidth(0.25)
    xlsXAxis.getFontFormat().setForeground (CLng(Red))

    'Modifying Y axis properties
    Set xlsYAxis = xlsChart.easy_getValueYAxis()
    xlsYAxis.getLineColorFormat().setLineColor (CLng(SteelBlue))
    xlsYAxis.getLineStyleFormat().setDashType (DASH_TYPE_LONG_DASH)
    xlsYAxis.getLineStyleFormat().setWidth(0.25)
    xlsYAxis.getFontFormat().setForeground (CLng(Blue))

    'Modifying series properties
    xlsChart.easy_getSeriesAt(0).getFillFormat().setBackground (CLng(RoyalBlue))
    xlsChart.easy_getSeriesAt(1).getFillFormat().setBackground (CLng(Yellow))
    xlsChart.easy_getSeriesAt(2).getFillFormat().setBackground (CLng(LightGreen))

    
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial23.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial23.xls")
    
    'Confirm generation
    dim sError
    sError = xls.easy_getError()
    if sError = "" then
		WScript.StdOut.Write(vbcrlf & "File successfully created. Press Enter to exit...")
    else
		WScript.StdOut.Write(vbcrlf & "Error: " & sError)
    end if
    WScript.StdIn.ReadLine()
    	
    'Dispose memory
	xls.Dispose
    