'----------------------------------------------------------------
' Tutorial 02
'
' This tutorial shows how to generate an Excel document from a list of values. 
' The cells are formatted using a predefined format.
'-----------------------------------------------------------------

Imports EasyXLS
Imports EasyXLS.Constants
Imports System.IO
Imports System.Data
Imports System.Drawing



Module Tutorial02

    Sub Main()


        Console.WriteLine("Tutorial 02" & vbCrLf & "----------" & vbCrLf)

        'Create an instance of the object that generates Excel files
        Dim xls As New ExcelDocument

        ' Create the database connection
        Dim sConnectionString As String = "Initial Catalog=Northwind;Data Source=localhost;User ID=sa;Password=;"
        Dim sqlConnection As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection(sConnectionString)
        sqlConnection.Open()

        ' Create the adapter used to fill the dataset
        Dim sQueryString As String = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', "
        sQueryString += " P.ProductName AS 'Product Name', O.UnitPrice AS Price, cast(O.Quantity AS varchar) As Quantity , O.UnitPrice * O. Quantity AS Value"
        sQueryString += " FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID"
        Dim adp As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(sQueryString, sqlConnection)

        ' Populate the dataset
        Dim ds As DataSet = New DataSet
        adp.Fill(ds)


        ' Create an instance of the object used to format the cells.
        Dim xlsAutoFormat As ExcelAutoFormat = New ExcelAutoFormat
        ' Set the style of the header
        Dim xlsHeaderStyle As ExcelStyle = New ExcelStyle(Color.LightGreen)
        xlsHeaderStyle.setFontSize(12)
        xlsAutoFormat.setHeaderRowStyle(xlsHeaderStyle)

        ' Set the style of the cells
        Dim xlsEvenRowStripesStyle As ExcelStyle = New ExcelStyle(Color.FloralWhite)
        xlsEvenRowStripesStyle.setFormat("$0.00")
        xlsEvenRowStripesStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)
        xlsAutoFormat.setEvenRowStripesStyle(xlsEvenRowStripesStyle)
        Dim xlsOddRowStripesStyle As ExcelStyle = New ExcelStyle(Color.FromArgb(240, 247, 239))
        xlsOddRowStripesStyle.setFormat("$0.00")
        xlsOddRowStripesStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)
        xlsAutoFormat.setOddRowStripesStyle(xlsOddRowStripesStyle)
        Dim xlsLeftColumnStyle As ExcelStyle = New ExcelStyle(Color.FloralWhite)
        xlsLeftColumnStyle.setFormat("mm/dd/yyyy")
        xlsLeftColumnStyle.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT)
        xlsAutoFormat.setLeftColumnStyle(xlsLeftColumnStyle)

        ' Generate the file
        Console.WriteLine("Writing file C:\\Samples\\Tutorial02.xls.")
        xls.easy_WriteXLSFile_FromDataSet("c:\\Samples\\Tutorial02.xls", ds, xlsAutoFormat, "Sheet1")

        ' Confirm generation
        Dim sError As String = xls.easy_getError()
        If (sError.Equals("")) Then
            Console.Write("File successfully created. Press Enter to Exit...")
        Else
            Console.Write("Error encountered: " + sError + "Press Enter to Exit...")
        End If


        ' Close the database connection.
        sqlConnection.Close()

        ' Dispose memory
        xls.Dispose()
        ds.Dispose()
        sqlConnection.Dispose()
        adp.Dispose()

        Console.ReadLine()

    End Sub

End Module
