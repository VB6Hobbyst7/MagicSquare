<?php
	/*==========================================================================
	 | Tutorial 10
	 |
	 | This tutorial shows how to merge a cell range.
	  ==========================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 10<br>";
	echo "----------<br>";
	


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
		
	//Add a worksheet
	$xls->easy_addWorksheet_2("Sheet1");

	//Get the table of the first sheet
	$xlsTable = $xls->easy_getSheet("Sheet1")->easy_getExcelTable();

	//Merging cells
	$xlsTable->easy_mergeCells_2("A1:C3");
	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial10.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial10.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>