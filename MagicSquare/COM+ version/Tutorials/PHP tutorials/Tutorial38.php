<?php
	/*==========================================================================
	 | Tutorial 38
	 |
	 | This tutorial shows how to load a XLSB file (we use the file
	 | generated in Tutorial 29), modify some data and save it to
	 | another file (Tutorial38.xlsb).
	  ==========================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 38<br>";
	echo "----------<br>";


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Read the file
	echo "Reading file: C:\\Samples\\Tutorial29.xlsb<br>";
	if ($xls->easy_LoadXLSBFile("C:\\Samples\\Tutorial29.xlsb"))
	{
		//Get the table of the second worksheet
		$xlsSecondTable = $xls->easy_getSheet("Second tab")->easy_getExcelTable();
		//Write some data
		$xlsSecondTable->easy_getCell_2("A1")->setValue("Data added by Tutorial38");

		for ($column=0; $column<5; $column++)
		{
			$xlsSecondTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		//Generate the file
		echo "Writing file: C:\Samples\Tutorial38.xlsb<br>";
		$xls->easy_WriteXLSBFile("C:\Samples\Tutorial38.xlsb");
		
		//Confirm generation
		if ($xls->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $xls->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial29.xlsb";
		echo $xls->easy_getError();
	}
	
	//Dispose memory
	$xls->Dispose();	
?>
