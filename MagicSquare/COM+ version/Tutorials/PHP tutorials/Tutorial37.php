<?php
	/*==========================================================================
	 | Tutorial 37
	 |
	 | This tutorial shows how to load a XLSX file (we use the file
	 | generated in Tutorial 28), modify some data and save it to
	 | another file (Tutorial37.xlsx).
	  ==========================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 37<br>";
	echo "----------<br>";


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Read the file
	echo "Reading file: C:\\Samples\\Tutorial28.xlsx<br>";
	if ($xls->easy_LoadXLSXFile("C:\\Samples\\Tutorial28.xlsx"))
	{
		//Get the table of the second worksheet
		$xlsSecondTable = $xls->easy_getSheet("Second tab")->easy_getExcelTable();
		//Write some data
		$xlsSecondTable->easy_getCell_2("A1")->setValue("Data added by Tutorial37");

		for ($column=0; $column<5; $column++)
		{
			$xlsSecondTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		//Generate the file
		echo "Writing file: C:\Samples\Tutorial37.xlsx<br>";
		$xls->easy_WriteXLSXFile("C:\Samples\Tutorial37.xlsx");
		
		//Confirm generation
		if ($xls->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $xls->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial28.xlsx";
		echo $xls->easy_getError();
	}
	
	//Dispose memory
	$xls->Dispose();	
?>
