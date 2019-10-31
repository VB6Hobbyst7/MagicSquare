<?php
	/*==========================================================================
	 | Tutorial 36
	 |
	 | This tutorial shows how to load an excel file (we use the one
	 | generated in Tutorial 09), modify some data and save it to
	 | another file (Tutorial36.xls).
	  ==========================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 36<br>";
	echo "----------<br>";
	

	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Read the file
	echo "Reading file: C:\\Samples\\Tutorial09.xls<br>";
	if ($xls->easy_LoadXLSFile("C:\\Samples\\Tutorial09.xls"))
	{
		//Get the table of the second worksheet
		$xlsSecondTable = $xls->easy_getSheet("Second tab")->easy_getExcelTable();

		//Write some data
		$xlsSecondTable->easy_getCell_2("A1")->setValue("Data added by Tutorial36");

		for ($column=0; $column<5; $column++)
		{
			$xlsSecondTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		//Generate the file
		echo "Writing file: C:\Samples\Tutorial36.xls<br>";
		$xls->easy_WriteXLSFile("C:\Samples\Tutorial36.xls");
		
		//Confirm generation
		if ($xls->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $xls->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial09.xls";
		echo $xls->easy_getError();
	}
	
	//Dispose memory
	$xls->Dispose();	
?>
