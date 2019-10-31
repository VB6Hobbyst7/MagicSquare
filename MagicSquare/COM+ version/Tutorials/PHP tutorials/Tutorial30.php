<?php
	/*==========================================================================
	 | Tutorial 30
	 |
	 | This tutorial shows how to export a CSV file.
	  ==========================================================================*/
	
	//Include files containing constants
	include("DataType.inc");

	header("Content-Type: text/html");

	echo "Tutorial 30<br>";
	echo "----------<br>";
	

	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Create the worksheet
	$xls->easy_addWorksheet_2("First tab");

	//Get the table of the first worksheet
	$xlsFirstTable = $xls->easy_getSheetAt(0)->easy_getExcelTable();

	//Add the cells for header
	for ($column=0; $column<5; $column++)
	{
		$xlsFirstTable->easy_getCell(0,$column)->setValue("Column " . ($column + 1));
		$xlsFirstTable->easy_getCell(0,$column)->setDataType($DATATYPE_STRING);
	}

	//Add the cells for data
	for ($row=0; $row<100; $row++)
	{
		for ($column=0; $column<5; $column++)
		{
			$xlsFirstTable->easy_getCell($row+1,$column)->setValue("Data ".($row + 1).", ".($column + 1));
			$xlsFirstTable->easy_getCell($row+1,$column)->setDataType($DATATYPE_STRING);
		}
	}

	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial30.csv<br>";
	$xls->easy_WriteCSVFile("C:\Samples\Tutorial30.csv","First tab");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>
