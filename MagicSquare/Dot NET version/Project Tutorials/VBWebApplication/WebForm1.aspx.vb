Imports EasyXLS
Imports EasyXLS.Charts
Imports EasyXLS.Constants

Public Class WebForm1
    Inherits System.Web.UI.Page



#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.dsSource = New System.Data.DataSet
        Me.dtTable = New System.Data.DataTable
        Me.DataColumn1 = New System.Data.DataColumn
        Me.DataColumn2 = New System.Data.DataColumn
        Me.DataColumn3 = New System.Data.DataColumn
        Me.DataColumn4 = New System.Data.DataColumn
        Me.DataColumn5 = New System.Data.DataColumn
        Me.DataColumn6 = New System.Data.DataColumn
        Me.DataColumn7 = New System.Data.DataColumn
        Me.DataColumn8 = New System.Data.DataColumn
        Me.DataColumn9 = New System.Data.DataColumn
        CType(Me.dsSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtTable, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'dsSource
        '
        Me.dsSource.DataSetName = "NewDataSet"
        Me.dsSource.Locale = New System.Globalization.CultureInfo("en-US")
        Me.dsSource.Tables.AddRange(New System.Data.DataTable() {Me.dtTable})
        '
        'dtTable
        '
        Me.dtTable.Columns.AddRange(New System.Data.DataColumn() {Me.DataColumn1, Me.DataColumn2, Me.DataColumn3, Me.DataColumn4, Me.DataColumn5, Me.DataColumn6, Me.DataColumn7, Me.DataColumn8, Me.DataColumn9})
        Me.dtTable.TableName = "dtTable"
        '
        'DataColumn1
        '
        Me.DataColumn1.ColumnName = "Project"
        '
        'DataColumn2
        '
        Me.DataColumn2.ColumnName = "Resource"
        '
        'DataColumn3
        '
        Me.DataColumn3.ColumnName = "Role"
        '
        'DataColumn4
        '
        Me.DataColumn4.ColumnName = "Task"
        '
        'DataColumn5
        '
        Me.DataColumn5.ColumnName = "Estimated"
        Me.DataColumn5.DataType = GetType(System.Int32)
        '
        'DataColumn6
        '
        Me.DataColumn6.ColumnName = "Regular"
        Me.DataColumn6.DataType = GetType(System.Int32)
        '
        'DataColumn7
        '
        Me.DataColumn7.ColumnName = "OT Hours"
        Me.DataColumn7.DataType = GetType(System.Int32)
        '
        'DataColumn8
        '
        Me.DataColumn8.ColumnName = "NB Hours"
        Me.DataColumn8.DataType = GetType(System.Int32)
        '
        'DataColumn9
        '
        Me.DataColumn9.ColumnName = "Approval Status"
        CType(Me.dsSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtTable, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents hlkEasyXLS As System.Web.UI.WebControls.HyperLink
    Protected WithEvents imgEasyXLSlogo As System.Web.UI.WebControls.Image
    Protected WithEvents btnExportToExcel As System.Web.UI.WebControls.Button
    Protected WithEvents dgTimeSheetReport As System.Web.UI.WebControls.DataGrid
    Protected WithEvents dsSource As System.Data.DataSet
    Protected WithEvents dtTable As System.Data.DataTable
    Protected WithEvents DataColumn1 As System.Data.DataColumn
    Protected WithEvents DataColumn2 As System.Data.DataColumn
    Protected WithEvents DataColumn3 As System.Data.DataColumn
    Protected WithEvents DataColumn4 As System.Data.DataColumn
    Protected WithEvents DataColumn5 As System.Data.DataColumn
    Protected WithEvents DataColumn6 As System.Data.DataColumn
    Protected WithEvents DataColumn7 As System.Data.DataColumn
    Protected WithEvents DataColumn8 As System.Data.DataColumn
    Protected WithEvents DataColumn9 As System.Data.DataColumn
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents chkNBHours As System.Web.UI.WebControls.CheckBox
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents chkTask As System.Web.UI.WebControls.CheckBox
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents chkEstimated As System.Web.UI.WebControls.CheckBox
    Protected WithEvents chkRegular As System.Web.UI.WebControls.CheckBox
    Protected WithEvents chkOTHours As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblError As System.Web.UI.WebControls.Label

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Populating the grid
        Dim Row1() As Object = {"EasyXLS", "Jim Bean", "Programmer", "Build Charts", 800, 240, 40, 0, "To be Approved"}
        dtTable.Rows.Add(Row1)
        Dim Row2() As Object = {"EasyXLS", "Jack White", "Programmer", "Build Worksheets", 1000, 160, 0, 0, "To be Approved"}
        dtTable.Rows.Add(Row2)
        Dim Row3() As Object = {"EasyXLS", "Christina Brown", "Programmer", "Build Hyperlinks", 750, 256, 2, 0, "To be Approved"}
        dtTable.Rows.Add(Row3)
        Dim Row4() As Object = {"EasyXLS", "Walt Whitman", "Programmer", "Create Tutorials", 600, 114, 10, 0, "To be Approved"}
        dtTable.Rows.Add(Row4)
        Dim Row5() As Object = {"EasyXLS", "Adam Wilson", "Tester", "Test Charts", 120, 8, 0, 0, "To be Approved"}
        dtTable.Rows.Add(Row5)
        Dim Row6() As Object = {"EasyXLS", "Will Crane", "Tester", "Test Hyperlinks", 100, 10, 2, 0, "To be Approved"}
        dtTable.Rows.Add(Row6)
        Dim Row7() As Object = {"EasyXLS", "George Brown", "Artist", "Design", 300, 150, 2, 0, "To be Approved"}
        dtTable.Rows.Add(Row7)
        Dim Row8() As Object = {"MS Excel", "Christian Wurm", "Programmer", "Database Design", 120, 35, 3, 0, "To be Approved"}
        dtTable.Rows.Add(Row8)
        Dim Row9() As Object = {"MS Excel", "Adrian Fisher", "Tester", "Speed", 240, 48, 0, 8, "To be Approved"}
        dtTable.Rows.Add(Row9)

        ' Computing the totals
        Dim nTotal As Integer
        nTotal = 0
        For nColumnIndex As Integer = 4 To 7
            nTotal = 0
            For nRowIndex As Integer = 0 To dtTable.Rows.Count - 1
                nTotal += Integer.Parse(dtTable.Rows(nRowIndex).ItemArray(nColumnIndex).ToString())
            Next
            dgTimeSheetReport.Columns(nColumnIndex).FooterText = nTotal.ToString()
        Next


        ' Data binding
        dgTimeSheetReport.DataBind()
    End Sub

    Private Sub btnExportToExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportToExcel.Click
        ' Creating an instance of the object that generates excel files
        Dim xls As New ExcelDocument

        ' Adding a sheet to the Excel Document object
        Dim xlsWorksheet As New ExcelWorksheet("TimeSheetReport")
        xls.easy_addWorksheet(xlsWorksheet)

        ' Adding the image
        xlsWorksheet.easy_addImage(Server.MapPath(imgEasyXLSlogo.ImageUrl), "A1")

        ' Adding the hyperlink
        xlsWorksheet.easy_addHyperlink(HyperlinkType.URL, hlkEasyXLS.NavigateUrl, "A5")


        ' Creating an instance of the object used to format the cells
        Dim xlsAutoFormat As New ExcelAutoFormat
        ' Setting the style of the header
        Dim xlsHeaderStyle As New ExcelStyle(dgTimeSheetReport.HeaderStyle.BackColor)
        xlsHeaderStyle.setBold(dgTimeSheetReport.HeaderStyle.Font.Bold)
        xlsAutoFormat.setHeaderRowStyle(xlsHeaderStyle)
        ' Setting the style of the cells
        xlsAutoFormat.setEvenRowStripesStyle(New ExcelStyle(dgTimeSheetReport.ItemStyle.BackColor))
        xlsAutoFormat.setOddRowStripesStyle(New ExcelStyle(dgTimeSheetReport.AlternatingItemStyle.BackColor))

        ' Adding the content of the grid
        xlsWorksheet.easy_insertDataSet(dsSource, 6, 0, xlsAutoFormat, True)

        ' Creating the footer
        Dim nFooterRowIndex As Integer = 6 + dtTable.Rows.Count + 1
        Dim xlsTable As ExcelTable = xlsWorksheet.easy_getExcelTable()
        xlsTable.easy_getCell(nFooterRowIndex, 0).setValue("Totals:")
        xlsTable.easy_getCell(nFooterRowIndex, 4).setValue("=SUM(E8:E" + nFooterRowIndex.ToString() + ")")
        xlsTable.easy_getCell(nFooterRowIndex, 5).setValue("=SUM(F8:F" + nFooterRowIndex.ToString() + ")")
        xlsTable.easy_getCell(nFooterRowIndex, 6).setValue("=SUM(G8:G" + nFooterRowIndex.ToString() + ")")
        xlsTable.easy_getCell(nFooterRowIndex, 7).setValue("=SUM(H8:H" + nFooterRowIndex.ToString() + ")")
        ' Setting the style of the footer
        Dim xlsFooterStyle As New ExcelStyle(dgTimeSheetReport.FooterStyle.BackColor)
        xlsFooterStyle.setBold(dgTimeSheetReport.FooterStyle.Font.Bold)
        xlsTable.easy_setRangeStyle(nFooterRowIndex, 0, nFooterRowIndex, 8, xlsFooterStyle)


        ' Creating and adding a chart based on the grid's data	
        Dim xlsChart As New ExcelChart("A20", 600, 300)
        If (chkEstimated.Checked) Then xlsChart.easy_addSeries("=TimeSheetReport!$E$7", "=TimeSheetReport!$E$8:$E$16")
        If (chkRegular.Checked) Then xlsChart.easy_addSeries("=TimeSheetReport!$F$7", "=TimeSheetReport!$F$8:$F$16")
        If (chkOTHours.Checked) Then xlsChart.easy_addSeries("=TimeSheetReport!$G$7", "=TimeSheetReport!$G$8:$G$16")
        If (chkNBHours.Checked) Then xlsChart.easy_addSeries("=TimeSheetReport!$H$7", "=TimeSheetReport!$H$8:$H$16")

        If (chkEstimated.Checked Or chkRegular.Checked Or chkOTHours.Checked Or chkNBHours.Checked) Then
            xlsChart.easy_setCategoryXAxisLabels("=TimeSheetReport!$D$8:$D$16")
        Else
            xlsChart.easy_addSeries("=TimeSheetReport!$D$7", "=TimeSheetReport!$D$8:$D$16")
        End If

        xlsWorksheet.easy_addChart(xlsChart)

        ' Preparing the Response object
        Response.AppendHeader("content-disposition", "attachment; filename=VBWebApplication.xls")
        Response.ContentType = "application/octetstream"
        Response.Clear()

        ' Generating the file and prompting the "Open or Save Dialog Box"
        Try
            xls.easy_WriteXLSFile(Response.OutputStream)
        Catch exc As Exception
            Response.ClearHeaders()
            Response.ClearContent()
            lblError.Text = exc.Message
        End Try

        xls.Dispose()

    End Sub
End Class
