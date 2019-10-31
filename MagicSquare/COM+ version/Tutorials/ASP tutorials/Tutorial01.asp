<%@ Language=VBScript %>
<!-- #INCLUDE FILE="Styles.inc" -->

<%
	'==========================================================================
	' Tutorial 01
	'
	' This tutorial shows how to generate an Excel document from a list of values. 
	' The cells are formatted using a predefined format.
	'==========================================================================

    
	response.write("Tutorial 01<br>")
	response.write("----------<br>")


	'Create an instance of the object that generates Excel files
	Set xls = Server.CreateObject("EasyXLS.ExcelDocument")

	'Connect to the database
	DIM objConn
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = "Provider=SQLOLEDB;Server=(local);Database=northwind;User ID=sa;Password=;"
	objConn.Open



	Dim sQueryString
	sQueryString = "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', P.ProductName AS 'Product Name', O.UnitPrice AS Price, cast(O.Quantity AS varchar) AS Quantity , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID"
	
	'Create the record set object
	Dim objRS
	Set objRS = Server.CreateObject("ADODB.Recordset") 
	objRS.Open sQueryString, objConn 
	
	'Create the list used to store the values
	Dim lstRows 
	Set lstRows = CreateObject("EasyXLS.Util.List")
	
	'Add the header row to the list
	Dim	 lstHeaderRow 	
	Set lstHeaderRow  = Server.CreateObject("EasyXLS.Util.List")
	lstHeaderRow.addElement("Order Date")
	lstHeaderRow.addElement("Product Name")
	lstHeaderRow.addElement("Price")
	lstHeaderRow.addElement("Quantity")
	lstHeaderRow.addElement("Value")	
	lstRows.addElement(lstHeaderRow)
	
	'Add the values from the database to the list
	Do Until objRS.EOF = True
		set RowList = Server.CreateObject("EasyXLS.Util.List")
		RowList.addElement("" & objRS("Order Date"))
		RowList.addElement("" & objRS("Product Name"))	
		RowList.addElement("" & objRS("Price"))
		RowList.addElement("" & objRS("Quantity"))
		RowList.addElement("" & objRS("Value"))
		lstRows.addElement(RowList)
			
	   'Move to the next record
	   objRS.MoveNext
	Loop 
	
	
	'Create an instance of the object used to format the cells
	Dim xlsAutoFormat 
	set xlsAutoFormat = Server.CreateObject("EasyXLS.ExcelAutoFormat")
	xlsAutoFormat.InitAs(AUTOFORMAT_EASYXLS1)
	
	
	'Generate the file
	response.write("Writing file: C:\Samples\Tutorial01.xls<br>")
	xls.easy_WriteXLSFile_FromList_2 "C:\Samples\Tutorial01.xls", lstRows, xlsAutoFormat, "Sheet1"

	
	'Confirm generation
	if xls.easy_getError() = "" then
		response.write("File successfully created.")
	else
		response.write("Error encountered: " + xls.easy_getError())
	end if
	
	'Close the Recordset object
	objRS.Close
	
	'Delete the Recordset Object
	Set objRS = Nothing
	
	
	'Close the Connection object
	objConn.Close
	
	'Delete the Connection Object
	Set objConn = Nothing 
	
	'Dispose memory
	xls.Dispose  
	

%>
