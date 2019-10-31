<?php
	/*==========================================================================
	 | Tutorial 32
	 |
	 | This tutorial shows how to export an XML file.
	  ==========================================================================*/
	
	//Include files containing constants
	include("DataType.inc");
	include("Styles.inc");

	header("Content-Type: text/html");

	echo "Tutorial 32<br>";
	echo "----------<br>";
	
    
	
	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Create the worksheets
	$xls->easy_addWorksheet_2("First tab");
	$xls->easy_addWorksheet_2("Second tab");

	//Get the table of the first worksheet
	$xlsFirstTable = $xls->easy_getSheetAt(0)->easy_getExcelTable();

	//Add the cells for header
	for ($column=0; $column<5; $column++)
	{
		$xlsFirstTable->easy_getCell(0,$column)->setValue("Column " . ($column + 1));
		$xlsFirstTable->easy_getCell(0,$column)->setDataType($DATATYPE_STRING);
	}
	$xlsFirstTable->easy_getRowAt(0)->setHeight(30);

	//Add the cells for data
	for ($row=0; $row<100; $row++)
	{
		for ($column=0; $column<5; $column++)
		{
			$xlsFirstTable->easy_getCell($row+1,$column)->setValue("Data ".($row + 1).", ".($column + 1));
			$xlsFirstTable->easy_getCell($row+1,$column)->setDataType($DATATYPE_STRING);
		}
	}

	// Create an instance of the object used to format the cells
	$xlsAutoFormat = new COM("EasyXLS.ExcelAutoFormat");
	$xlsAutoFormat->InitAs($AUTOFORMAT_EASYXLS1);

	//Apply the predefined format to the cells.
	$xlsFirstTable->easy_setRangeAutoFormat_2("A1:E101", $xlsAutoFormat);

	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial32.xml<br>";
	$xls->easy_WriteXMLFile_2("C:\Samples\Tutorial32.xml");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>
