<?php
	/*==========================================================================
	 | Tutorial 27
	 |
	 | This tutorial shows how to encrypt and set the password required for opening a document
	  ==========================================================================*/

	header("Content-Type: text/html");
	
	echo "Tutorial 27<br>";
	echo "----------<br>";
	


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	

	//Create the worksheets 
	$xls->easy_addWorksheet_2("First tab");
	$xls->easy_addWorksheet_2("Second tab");

	//Set the password required for opening the document
	$xls->easy_getOptions()->setPasswordToOpen("password");
		
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial27.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial27.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>
