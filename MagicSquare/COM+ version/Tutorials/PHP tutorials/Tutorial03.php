<?php
	/*==========================================================================
	 | Tutorial 03
	 |
	 | This tutorial shows how to create a Microsoft Excel file
	 | that has two worksheets.
	  ==========================================================================*/

	header("Content-Type: text/html");
	
	echo "Tutorial 03<br>";
	echo "----------<br>";
	


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	

	//Create the worksheets 
	$xls->easy_addWorksheet_2("First tab");
	$xls->easy_addWorksheet_2("Second tab");

	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial03.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial03.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>