<?php
	/*==========================================================================
	 | Tutorial 11
	 |
	 | This tutorial shows how to create a Microsoft Excel file
	 | that has a formula.
	  ==========================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 11<br>";
	echo "----------<br>";
	


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Add one worksheet
	$xls->easy_addWorksheet_2("Formula");
		
	//Get the table, populate the sheet and set a formula
	$xlsTable = $xls->easy_getSheet("Formula")->easy_getExcelTable();
	$xlsTable->easy_getCell_2("A1")->setValue("1");
	$xlsTable->easy_getCell_2("A2")->setValue("2");
	$xlsTable->easy_getCell_2("A3")->setValue("3");
	$xlsTable->easy_getCell_2("A4")->setValue("4");
	$xlsTable->easy_getCell_2("A6")->setValue("=SUM(A1:A4)");
	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial11.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial11.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>