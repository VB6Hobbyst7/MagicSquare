<?php
	/*==========================================================================
	 | Tutorial 39
	 |
	 | This tutorial shows how to load a CSV file (we use the file
	 | generated in Tutorial 30), modify some data and save it to
	 | another file (Tutorial39.xls).
	  ==========================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 39<br>";
	echo "----------<br>";


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Read the file
	echo "Reading file: C:\\Samples\\Tutorial30.csv<br>";
	if ($xls->easy_LoadCSVFile("C:\\Samples\\Tutorial30.csv"))
	{
		
		//Set the name of the first worksheet
		$xls->easy_getSheetAt(0)->setSheetName("First tab");

		//Add a new worksheet and write some data
		$xls->easy_addWorksheet_2("Second tab");
		$xlsTable = $xls->easy_getSheetAt(1)->easy_getExcelTable();
		$xlsTable->easy_getCell_2("A1")->setValue("Data added by Tutorial39");

		for ($column=0; $column<5; $column++)
		{
			$xlsTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		//Generate the file
		echo "Writing file: C:\Samples\Tutorial39.xls<br>";
		$xls->easy_WriteXLSFile("C:\Samples\Tutorial39.xls");
		
		//Confirm generation
		if ($xls->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $xls->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial30.csv";
		echo $xls->easy_getError();
	}
	
	//Dispose memory
	$xls->Dispose();	
?>
