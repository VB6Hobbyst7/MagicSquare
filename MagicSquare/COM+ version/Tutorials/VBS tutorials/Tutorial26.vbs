    '==========================================================================
    'Tutorial 26
    '
    ' This tutorial shows how to create a pivot chart. The pivot chart is
    ' added to a workshet and also to a separate chart sheet.
    '==========================================================================
    
    WScript.StdOut.WriteLine("Tutorial 26" & vbcrlf & "----------" & vbcrlf)
    
	'Constants declaration
    Dim DATATYPE_STRING
    DATATYPE_STRING = "string"
	Dim PIVOTTABLE_SUBTOTAL_SUM
    PIVOTTABLE_SUBTOTAL_SUM = 2
	Dim PIVOTTABLE_PIVOT_STYLE_MEDIUM_9 
    PIVOTTABLE_PIVOT_STYLE_MEDIUM_9 = "PivotStyleMedium9"
	Dim CHART_CHART_TYPE_PYRAMID_BAR 
    CHART_CHART_TYPE_PYRAMID_BAR = 133
	

    
   
    'Create an instance of the object that generates Excel files
    Set xls = CreateObject("EasyXLS.ExcelDocument")
    
	'Create the worksheets
	xls.easy_addWorksheet_2("First tab")
	xls.easy_addWorksheet_2("Second tab")

	'Create the chart sheet
	xls.easy_addChart_2("Pivot chart")
    
	'Get the table of the first worksheet
	Set xlsFirstTable = xls.easy_getSheetAt(0).easy_getExcelTable()
	
	'Add the cells for header
	xlsFirstTable.easy_getCell(0,0).setValue("Sale agent")
	xlsFirstTable.easy_getCell(0,0).setDataType(DATATYPE_STRING)
	xlsFirstTable.easy_getCell(0,1).setValue("Sale country")
	xlsFirstTable.easy_getCell(0,1).setDataType(DATATYPE_STRING)
	xlsFirstTable.easy_getCell(0,2).setValue("Month")
	xlsFirstTable.easy_getCell(0,2).setDataType(DATATYPE_STRING)
	xlsFirstTable.easy_getCell(0,3).setValue("Year")
	xlsFirstTable.easy_getCell(0,3).setDataType(DATATYPE_STRING)
	xlsFirstTable.easy_getCell(0,4).setValue("Sale amount")
	xlsFirstTable.easy_getCell(0,4).setDataType(DATATYPE_STRING)

	xlsFirstTable.easy_getRowAt(0).setBold(True)

	'Populate the source for pivot table
	xlsFirstTable.easy_getCell(1,0).setValue("John Down")
	xlsFirstTable.easy_getCell(1,1).setValue("USA")
	xlsFirstTable.easy_getCell(1,2).setValue("June")
	xlsFirstTable.easy_getCell(1,3).setValue("2010")
	xlsFirstTable.easy_getCell(1,4).setValue("550")
		
	xlsFirstTable.easy_getCell(2,0).setValue("Scott Valey")
	xlsFirstTable.easy_getCell(2,1).setValue("United Kingdom")
	xlsFirstTable.easy_getCell(2,2).setValue("June")
	xlsFirstTable.easy_getCell(2,3).setValue("2010")
	xlsFirstTable.easy_getCell(2,4).setValue("2300")
		
	xlsFirstTable.easy_getCell(3,0).setValue("John Down")
	xlsFirstTable.easy_getCell(3,1).setValue("USA")
	xlsFirstTable.easy_getCell(3,2).setValue("July")
	xlsFirstTable.easy_getCell(3,3).setValue("2010")
	xlsFirstTable.easy_getCell(3,4).setValue("3100")
		
	xlsFirstTable.easy_getCell(4,0).setValue("John Down")
	xlsFirstTable.easy_getCell(4,1).setValue("USA")
	xlsFirstTable.easy_getCell(4,2).setValue("June")
	xlsFirstTable.easy_getCell(4,3).setValue("2011")
	xlsFirstTable.easy_getCell(4,4).setValue("1050")
			
	xlsFirstTable.easy_getCell(5,0).setValue("John Down")
	xlsFirstTable.easy_getCell(5,1).setValue("USA")
	xlsFirstTable.easy_getCell(5,2).setValue("July")
	xlsFirstTable.easy_getCell(5,3).setValue("2011")
	xlsFirstTable.easy_getCell(5,4).setValue("2400")
		
	xlsFirstTable.easy_getCell(6,0).setValue("Steve Marlowe")
	xlsFirstTable.easy_getCell(6,1).setValue("France")
	xlsFirstTable.easy_getCell(6,2).setValue("June")
	xlsFirstTable.easy_getCell(6,3).setValue("2011")
	xlsFirstTable.easy_getCell(6,4).setValue("1200")
		
	xlsFirstTable.easy_getCell(7,0).setValue("Scott Valey")
	xlsFirstTable.easy_getCell(7,1).setValue("United Kingdom")
	xlsFirstTable.easy_getCell(7,2).setValue("June")
	xlsFirstTable.easy_getCell(7,3).setValue("2011")
	xlsFirstTable.easy_getCell(7,4).setValue("700")
		
	xlsFirstTable.easy_getCell(8,0).setValue("Scott Valey")
	xlsFirstTable.easy_getCell(8,1).setValue("United Kingdom")
	xlsFirstTable.easy_getCell(8,2).setValue("July")
	xlsFirstTable.easy_getCell(8,3).setValue("2011")
	xlsFirstTable.easy_getCell(8,4).setValue("360")
		
	'Create pivot table
	Set xlsPivotTable = CreateObject("EasyXLS.PivotTables.ExcelPivotTable")
		
	xlsPivotTable.setName("Sales")
	xlsPivotTable.setSourceRange "First tab!$A$1:$E$9", xls
	xlsPivotTable.setLocation_2("A3:G15")
	xlsPivotTable.addFieldToRowLabels("Sale agent")
	xlsPivotTable.addFieldToColumnLabels("Year")
	xlsPivotTable.addFieldToValues "Sale amount","Sale amount per year",PIVOTTABLE_SUBTOTAL_SUM 
	xlsPivotTable.addFieldToReportFilter("Sale country")
	xlsPivotTable.setOutlineForm()
	xlsPivotTable.setStyle(PIVOTTABLE_PIVOT_STYLE_MEDIUM_9)

	'Add the pivot table
	Set xlsWorksheet = xls.easy_getSheet("Second tab")
	xlsWorksheet.easy_addPivotTable(xlsPivotTable)

	'Create a pivot chart
        set xlsPivotChart1 = CreateObject("EasyXLS.PivotTables.ExcelPivotChart")
        xlsPivotChart1.setSize 600, 300
        xlsPivotChart1.setLeftUpperCorner_2("A10")
        xlsPivotChart1.easy_setChartType(CHART_CHART_TYPE_PYRAMID_BAR)
        xlsPivotChart1.getChartTitle().setText("Sales")
        xlsPivotChart1.setPivotTable(xlsPivotTable)

        'Add the pivot chart to the worksheet
        xlsWorksheet.easy_addPivotChart(xlsPivotChart1)

	'Create a clone of the pivot chart and add the clone to the chart sheet
        set xlsPivotChart2 = xlsPivotChart1.Clone()
        xlsPivotChart2.setSize 970, 630
	set xlsChartSheet = xls.easy_getSheet("Pivot chart")
        xlsChartSheet.easy_setExcelChart(xlsPivotChart2)
			
	
    	'Generate the file
    	WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial26.xlsx")
    	xls.easy_WriteXLSXFile ("C:\Samples\Tutorial26.xlsx")
    
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