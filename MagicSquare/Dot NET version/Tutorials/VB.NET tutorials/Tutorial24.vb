'----------------------------------------------------------------
' Tutorial 24
'
' This tutorial shows how to create and insert a chart in a worksheet.
'-----------------------------------------------------------------*/

Imports EasyXLS
Imports EasyXLS.Charts
Imports EasyXLS.Constants

Module Tutorial24

    Sub Main()



        Console.WriteLine("Tutorial 24" & vbCrLf & "----------" & vbCrLf)

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

        'Create the chart.
        Dim xlsChart As ExcelChart = New ExcelChart("A10", 600, 300)
        xlsChart.easy_addSeries("=SourceData!$B$1", "=SourceData!$B$2:$B$5")
        xlsChart.easy_addSeries("=SourceData!$C$1", "=SourceData!$C$2:$C$5")
        xlsChart.easy_addSeries("=SourceData!$D$1", "=SourceData!$D$2:$D$5")
        xlsChart.easy_setCategoryXAxisLabels("=SourceData!$A$2:$A$5")

        'Add the chart to the first worksheet.
        Dim xlsWorksheet As ExcelWorksheet = xls.easy_getSheet("SourceData")
        xlsWorksheet.easy_addChart(xlsChart)

        'Generate the file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial24.xls.")
        xls.easy_WriteXLSFile("C:\\Samples\\Tutorial24.xls")

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
