<?php
	/*==========================================================================
	 | Tutorial 02
	 |
	 | This tutorial shows how to generate an Excel document from a list of values. 
	 | The cells are formatted using an user-defined format.
	  ==========================================================================*/
  	//Include Files
	include("Color.inc");
	include("Alignment.inc");

    // Constants declaration
    $OddRowStripesStyleColor = 0xfff0f7ef;
    
	header("Content-Type: text/html");
	
	echo "Tutorial 02<br>";
	echo "----------<br>";
		

	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");


	// Connect to the database
	$serverName = "(local)";
	$connectionInfo = array("Database"=>"northwind","UID"=>"sa","PWD"=>"");
	
	$db_conn = sqlsrv_connect( $serverName, $connectionInfo); 
	if( $db_conn === false )  
	{
   	  echo "Unable to connect.";
  	   die( print_r( sqlsrv_errors(), true));
	}

	
	// Query the database
	$query_result = sqlsrv_query( $db_conn, "SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', P.ProductName AS 'Product Name', O.UnitPrice AS Price, ' ' + cast(O.Quantity AS varchar) AS Quantity , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID" )
		or die( "<strong>ERROR: Query failed</strong>" );


	// Create the list used to store the values
	$lstRows = new COM("EasyXLS.Util.List");
	
	// Add the header row to the list
	$lstHeaderRow  = new COM("EasyXLS.Util.List");
	$lstHeaderRow->addElement("Order Date");
	$lstHeaderRow->addElement("Product Name");
	$lstHeaderRow->addElement("Price");
	$lstHeaderRow->addElement("Quantity");
	$lstHeaderRow->addElement("Value");
	$lstRows->addElement($lstHeaderRow);
	
		
	// Add the values from the database to the list
	while ($row=sqlsrv_fetch_array($query_result))
	{
		$RowList = new COM("EasyXLS.Util.List");
		$RowList->addElement("" . $row['Order Date']);
		$RowList->addElement("" . $row["Product Name"]);
		$RowList->addElement("" . $row["Price"]);
		$RowList->addElement("" . $row["Quantity"]);
		$RowList->addElement("" . $row["Value"]);
		$lstRows->addElement($RowList);
			
	}
	
	// Create an instance of the object used to format the cells
	$xlsAutoFormat = new COM("EasyXLS.ExcelAutoFormat");
	//Set the style of the header
	$xlsHeaderStyle = new COM("EasyXLS.ExcelStyle");
	$xlsHeaderStyle->setBackground((int)$COLOR_LIGHTGREEN);
	$xlsHeaderStyle->setFontSize(12);
	$xlsAutoFormat->setHeaderRowStyle($xlsHeaderStyle);

	// Set the style of the cells
	$xlsEvenRowStripesStyle = new COM("EasyXLS.ExcelStyle");
	$xlsEvenRowStripesStyle->setBackground((int)$COLOR_FLORALWHITE);
	$xlsEvenRowStripesStyle->setFormat("$0.00");
	$xlsEvenRowStripesStyle->setHorizontalAlignment($ALIGNMENT_ALIGNMENT_LEFT);
	$xlsAutoFormat->setEvenRowStripesStyle($xlsEvenRowStripesStyle)	;
	$xlsOddRowStripesStyle = new COM("EasyXLS.ExcelStyle");
	$xlsOddRowStripesStyle->setBackground((int)$OddRowStripesStyleColor);
	$xlsOddRowStripesStyle->setFormat("$0.00");
	$xlsOddRowStripesStyle->setHorizontalAlignment ($ALIGNMENT_ALIGNMENT_LEFT);
	$xlsAutoFormat->setOddRowStripesStyle($xlsOddRowStripesStyle);
	$xlsLeftColumnStyle = new COM("EasyXLS.ExcelStyle");
	$xlsLeftColumnStyle->setBackground((int)$COLOR_FLORALWHITE);
	$xlsLeftColumnStyle->setFormat("mm/dd/yyyy");
	$xlsLeftColumnStyle->setHorizontalAlignment($ALIGNMENT_ALIGNMENT_LEFT);
	$xlsAutoFormat->setLeftColumnStyle($xlsLeftColumnStyle);
	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial02.xls<br>";
	$xls->easy_WriteXLSFile_FromList_2("C:\Samples\Tutorial02.xls", $lstRows, $xlsAutoFormat, "Sheet1");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		

	// Free the memory associated with the query
	sqlsrv_free_stmt( $query_result );

	// Close the Connection object
	sqlsrv_close( $db_conn );           
  
  	//Dispose memory
	$xls->Dispose();

?>