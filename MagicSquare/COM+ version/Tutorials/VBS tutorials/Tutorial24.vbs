    '==========================================================================
    ' Tutorial 24
    '
    ' This tutorial shows how to create and insert a chart in a worksheet.
    '==========================================================================
  
	'Constants declaration
    Dim FORMAT_DATE
    FORMAT_DATE = "MM/dd/yyyy"

    
     WScript.StdOut.WriteLine("Tutorial 24" & vbcrlf & "-----------" & vbcrlf)
    
   

   
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
      
   	
	'Create the chart.
	set xlsChart = CreateObject("EasyXLS.Charts.ExcelChart")
	xlsChart.setLeftUpperCorner_2("A10")
	xlsChart.setSize 600, 300
		
	xlsChart.easy_addSeries_2 "=SourceData!$B$1", "=SourceData!$B$2:$B$5"
	xlsChart.easy_addSeries_2 "=SourceData!$C$1", "=SourceData!$C$2:$C$5"
	xlsChart.easy_addSeries_2 "=SourceData!$D$1", "=SourceData!$D$2:$D$5"
	xlsChart.easy_setCategoryXAxisLabels("=SourceData!$A$2:$A$5")

	'Add the chart to the first worksheet.
	set xlsWorksheet = xls.easy_getSheet("SourceData")
	xlsWorksheet.easy_addChart(xlsChart)
    
    'Generate the file
    WScript.StdOut.WriteLine("Writing file C:\Samples\Tutorial24.xls")
    xls.easy_WriteXLSFile ("C:\Samples\Tutorial24.xls")
    
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