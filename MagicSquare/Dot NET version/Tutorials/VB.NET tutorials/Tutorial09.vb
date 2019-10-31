'----------------------------------------------------------------
' Tutorial 09
'
' This tutorial shows how to create a Microsoft Excel file
' that has two worksheets. The first one is full with data
' and the cells are formatted. The column header has comments.
' The first worksheet has header & footer. The print options are
' set for the first worksheet.
'-----------------------------------------------------------------

Imports System.Drawing
Imports EasyXLS
Imports EasyXLS.Constants

Module Tutorial09

    Sub Main()


        Console.WriteLine("Tutorial 09" & vbCrLf & "----------" & vbCrLf)


        'Create an instance of the object that generates Excel files, having 2 sheets
        Dim xls As New ExcelDocument(2)

        'Set the sheet names
        xls.easy_getSheetAt(0).setSheetName("First tab")
        xls.easy_getSheetAt(1).setSheetName("Second tab")

        'Lock the first tab
        xls.easy_getSheetAt(0).setSheetProtected(True)

        'Get the table of the first worksheet
        Dim xlsFirstTab As ExcelWorksheet = xls.easy_getSheetAt(0)
        Dim xlsFirstTable = xlsFirstTab.easy_getExcelTable()

        'Create the style for the header
        Dim xlsStyleHeader As New ExcelStyle("Verdana", 8, True, True, Color.Yellow)
        xlsStyleHeader.setBackground(Color.Black)
        xlsStyleHeader.setBorderColors(Color.Gray, Color.Gray, Color.Gray, Color.Gray)
        xlsStyleHeader.setBorderStyles(Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM)
        xlsStyleHeader.setHorizontalAlignment(Alignment.ALIGNMENT_CENTER)
        xlsStyleHeader.setVerticalAlignment(Alignment.ALIGNMENT_BOTTOM)
        xlsStyleHeader.setWrap(True)
        xlsStyleHeader.setDataType(DataType.STRING)

        'Add the cells for header
        For column As Integer = 0 To 4
            xlsFirstTable.easy_getCell(0, column).setValue("Column " & (column + 1))
            xlsFirstTable.easy_getCell(0, column).setStyle(xlsStyleHeader)

            'Add comment
            xlsFirstTable.easy_getCell(0, column).setComment("This is column no " & (column + 1))
        Next
        xlsFirstTable.easy_getRowAt(0).setHeight(30)


        'Add the cells for data
        For row As Integer = 0 To 99
            For column As Integer = 0 To 4
                xlsFirstTable.easy_getCell(row + 1, column).setValue("Data " & (row + 1) & ", " & (column + 1))
            Next
        Next

        'Create a style for cells
        Dim xlsStyleData As New ExcelStyle
        xlsStyleData.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)
        xlsStyleData.setForeground(Color.DarkGray)
        xlsStyleData.setWrap(False)
        xlsStyleData.setDataType(DataType.STRING)
        xlsStyleData.setLocked(True)
        xlsFirstTable.easy_setRangeStyle("A2:E101", xlsStyleData)

        'Set column widths
        xlsFirstTable.setColumnWidth(0, 70)
        xlsFirstTable.setColumnWidth(1, 100)
        xlsFirstTable.setColumnWidth(2, 70)
        xlsFirstTable.setColumnWidth(3, 100)
        xlsFirstTable.setColumnWidth(4, 70)

        'Add headers for the first worksheet
        xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertSingleUnderline()
        xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertFile()
        xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertValue(" - How to create header and footer")

        xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertDate()
        xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertValue(" ")
        xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertTime()

        'Add footer for the first worksheet
        xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertPage()
        xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertValue(" of ")
        xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertPages()

        'Set Page Setup options
        Dim xlsPageSetup = xlsFirstTab.easy_getPageSetup()
        xlsPageSetup.easy_setPrintArea("A1:E101")
        xlsPageSetup.easy_setRowsToRepeatAtTop("$1:$1")
        xlsPageSetup.setCenterHorizontally(True)
        xlsPageSetup.setOrientation(PageSetup.ORIENTATION_PORTRAIT)
        xlsPageSetup.setPageOrder(PageSetup.PAGE_ORDER_DOWN_THEN_OVER)
        xlsPageSetup.setPaperSize(PageSetup.PAPER_SIZE_A4)
        xlsPageSetup.setPrintComments(PageSetup.COMMENTS_AT_END_OF_SHEET)
        xlsPageSetup.setPrintGridlines(True)
        xlsFirstTable.easy_insertPageBreakAtRow(21)
        xlsFirstTable.easy_insertPageBreakAtRow(41)
        xlsFirstTable.easy_insertPageBreakAtRow(61)
        xlsFirstTable.easy_insertPageBreakAtRow(81)
        xlsFirstTab.setPageBreakPreview(True)


        'Generate the file
        Console.WriteLine("Writing file C:\Samples\Tutorial09.xls.")
        xls.easy_WriteXLSFile("C:\Samples\Tutorial09.xls")

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
