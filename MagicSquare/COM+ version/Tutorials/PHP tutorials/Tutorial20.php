<?php
	/*==========================================================================
	 | Tutorial 20
	 |
	 | This tutorial shows how to create a Microsoft Excel file 
 	 | that has AutoFilter.

	  ==========================================================================*/

	//Include files containing constants
	include("DataType.inc");

	header("Content-Type: text/html");

	echo "Tutorial 20<br>";
	echo "----------<br>";

	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");

	//Create the worksheets
	$xls->easy_addWorksheet_2("Sheet1");

	//Get the table of the first worksheet
	$xlsTab = $xls->easy_getSheet("Sheet1");
	$xlsTable = $xlsTab->easy_getExcelTable();

	//Add the cells for header
	for ($column=0; $column<5; $column++)
	{
		$xlsTable->easy_getCell(0,$column)->setValue("Column " . ($column + 1));
		$xlsTable->easy_getCell(0,$column)->setDataType($DATATYPE_STRING);
	}
	
	//Add the cells for data
	for ($row=0; $row<100; $row++)
	{
		for ($column=0; $column<5; $column++)
		{
			$xlsTable->easy_getCell($row+1,$column)->setValue("Data ".($row + 1).", ".($column + 1));
			$xlsTable->easy_getCell($row+1,$column)->setDataType($DATATYPE_STRING);
		}
	}
	
	//Add AutoFilter
	$xlsFilter = $xlsTab->easy_getFilter();
   	$xlsFilter->setAutoFilter_2("A1:E1");
	

	//Generate the file
	echo "Writing file: C:\Samples\Tutorial20.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial20.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>