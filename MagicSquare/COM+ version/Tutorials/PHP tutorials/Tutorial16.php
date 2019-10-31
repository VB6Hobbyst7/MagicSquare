<?php
	/*==========================================================================
	 | Tutorial 16
	 |
	 | This tutorial shows how to create a Microsoft Excel file
	 | that has two worksheets. The first one has an image.
	  ==========================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 16<br>";
	echo "----------<br>";
	


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Create the worksheets 
	$xls->easy_addWorksheet_2("First tab");
	$xls->easy_addWorksheet_2("Second tab");

	//Create the image
	$xls->easy_getSheetAt(0)->easy_addImage_5("C:\\Samples\\EasyXLSLogo.JPG", "A1");

	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial16.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial16.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>