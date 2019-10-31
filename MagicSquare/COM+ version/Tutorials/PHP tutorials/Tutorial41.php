<?php
	/*==========================================================================
	 | Tutorial 41
	 |
	 | This tutorial shows how to load an XML file (we use the file
	 | generated in Tutorial 32), modify some data and save it to
	 | another file (Tutorial41.xls).
	  ==========================================================================*/
	
	header("Content-Type: text/html");

	echo "Tutorial 41<br>";
	echo "----------<br>";



	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Read the file
	echo "Reading file: C:\\Samples\\Tutorial32.xml<br>";
	if ($xls->easy_LoadXMLSpreadsheetFile_2("C:\\Samples\\Tutorial32.xml"))
	{
		//Get the table of the second worksheet and write some data
		$xlsTable = $xls->easy_getSheetAt(1)->easy_getExcelTable();
		$xlsTable->easy_getCell_2("A1")->setValue("Data added by Tutorial41");

		for ($column=0; $column<5; $column++)
		{
			$xlsTable->easy_getCell(1, $column)->setValue("Data " . ($column + 1));
		}
		
		//Generate the file
		echo "Writing file: C:\Samples\Tutorial41.xls<br>";
		$xls->easy_WriteXLSFile("C:\Samples\Tutorial41.xls");
		
		//Confirm generation
		if ($xls->easy_getError() == "")
			echo "File successfully created.";
		else
			echo "Error encountered: " . $xls->easy_getError();
	}
	else
	{
		echo "Error reading file C:\Samples\Tutorial32.xml";
		echo $xls->easy_getError();
	}
	
	//Dispose memory
	$xls->Dispose();	
?>
