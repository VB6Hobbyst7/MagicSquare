'----------------------------------------------------------------
' Tutorial 22
'
' This tutorial shows how to show the chart data table and
' to set its properties.
'-----------------------------------------------------------------

Imports System.Drawing
Imports EasyXLS
Imports EasyXLS.Charts
Imports EasyXLS.Constants
Imports EasyXLS.Drawings.Formatting


Module Tutorial22

    Sub Main()



        Console.WriteLine("Tutorial 22" & vbCrLf & "----------" & vbCrLf)

        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument

        'Add one worksheet
        xls.easy_addWorksheet("SourceData")

        ' ----------------------------------------------------------------------
        'Insert values
        Dim xlsFirstTab As ExcelWorksheet = xls.easy_getSheet("SourceData")
        Dim xlsTable1 = xlsFirstTab.easy_getExcelTable()

        xlsTable1.easy_getCell(0, 0).setValue("Show Date")
        xlsTable1.easy_getCell(0, 1).setValue("Available Places")
        xlsTable1.easy_getCell(0, 2).setValue("Available Tickets")
        xlsTable1.easy_getCell(0, 3).setValue("Sold Tickets")

        xlsTable1.easy_getCell(1, 0).setValue("03/13/2005 00:00:00")
        xlsTable1.easy_getCell(1, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE)
        xlsTable1.easy_getCell(2, 0).setValue("03/14/2005 00:00:00")
        xlsTable1.easy_getCell(2, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE)
        xlsTable1.easy_getCell(3, 0).setValue("03/15/2005 00:00:00")
        xlsTable1.easy_getCell(3, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE)
        xlsTable1.easy_getCell(4, 0).setValue("03/16/2005 00:00:00")
        xlsTable1.easy_getCell(4, 0).setFormat(EasyXLS.Constants.Format.FORMAT_DATE)

        xlsTable1.easy_getCell(1, 1).setValue("10000")
        xlsTable1.easy_getCell(2, 1).setValue("5000")
        xlsTable1.easy_getCell(3, 1).setValue("8500")
        xlsTable1.easy_getCell(4, 1).setValue("1000")

        xlsTable1.easy_getCell(1, 2).setValue("8000")
        xlsTable1.easy_getCell(2, 2).setValue("4000")
        xlsTable1.easy_getCell(3, 2).setValue("6000")
        xlsTable1.easy_getCell(4, 2).setValue("1000")

        xlsTable1.easy_getCell(1, 3).setValue("920")
        xlsTable1.easy_getCell(2, 3).setValue("1005")
        xlsTable1.easy_getCell(3, 3).setValue("342")
        xlsTable1.easy_getCell(4, 3).setValue("967")

        xlsTable1.easy_getColumnAt(0).setWidth(100)
        xlsTable1.easy_getColumnAt(1).setWidth(100)
        xlsTable1.easy_getColumnAt(2).setWidth(100)
        xlsTable1.easy_getColumnAt(3).setWidth(100)
        '--------------------------------------------------------------------------

        'Add the chart
        xls.easy_addChart("Chart", "=SourceData!$A$1:$D$5", Chart.SERIES_IN_COLUMNS)

        'Get the previously added chart
        Dim xlsChartSheet As ExcelChartSheet = xls.easy_getSheetAt(1)
        Dim xlsChart As ExcelChart = xlsChartSheet.easy_getExcelChart()


        'Hiding the legend
        xlsChart.easy_getLegend().setVisible(False)

        'Make DataTable visible
        xlsChart.easy_getChartDataTable().setVisible(True)
        xlsChart.easy_getChartDataTable().getFontFormat().setFont("Verdana")
        xlsChart.easy_getChartDataTable().getFontFormat().setFontSize(10)
        xlsChart.easy_getChartDataTable().setHorizontalLines(False)
        xlsChart.easy_getChartDataTable().setLegendKey(True)
        xlsChart.easy_getChartDataTable().getLineColorFormat().setLineColor(Color.Blue)
        xlsChart.easy_getChartDataTable().setVerticalLines(False)


        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial22.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial22.xls")

        'Confirm generation
        Dim sError As String = xls.easy_getError()
        If (sError.Equals("")) Then
            Console.Write(vbCrLf & "File successfully created. Press Enter to Exit...")
        Else
            Console.Write(vbCrLf & "Error encountered: " & sError & vbCrLf & "Press Enter to Exit...")
        End If
        Console.ReadLine()

    End Sub

End Module
