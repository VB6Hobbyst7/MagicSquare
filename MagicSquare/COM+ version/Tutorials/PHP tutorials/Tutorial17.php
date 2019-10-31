<?php
	/*==========================================================================
	 | Tutorial 17
	 |
	 | This tutorial shows how to create a Microsoft Excel file
	 | that has two worksheets. The first one is full with data
	 | and contains groups.
	  ==========================================================================*/
	
	//Include files containing constants
	include("DataType.inc");
	include("Styles.inc");
	include("DataGroup.inc");



	header("Content-Type: text/html");
	
	echo "Tutorial 17<br>";
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
	for ($row=0; $row<25; $row++)
	{
		for ($column=0; $column<5; $column++)
		{
			$xlsFirstTable->easy_getCell($row+1,$column)->setValue("Data ".($row + 1).", ".($column + 1));
			$xlsFirstTable->easy_getCell($row+1,$column)->setDataType($DATATYPE_STRING);
		}
	}

	//Set column widths
	$xlsFirstTable->setColumnWidth_2(0, 70);
	$xlsFirstTable->setColumnWidth_2(1, 100);
	$xlsFirstTable->setColumnWidth_2(2, 70);
	$xlsFirstTable->setColumnWidth_2(3, 100);
	$xlsFirstTable->setColumnWidth_2(4, 70);
	
	
	//Create the first group
    $xlsFirstDataGroup = new COM("EasyXLS.ExcelDataGroup");
    $xlsFirstDataGroup->setRange_2 ("A1:E26");
    $xlsFirstDataGroup->setGroupType ($DATAGROUP_GROUP_BY_ROWS);
    $xlsFirstDataGroup->setCollapsed (False);
    
    //Create an instance of the object used to format the cells of the first group
    $xlsAutoFormat = new COM("EasyXLS.ExcelAutoFormat");
    $xlsAutoFormat->InitAs($AUTOFORMAT_EASYXLS1);
    $xlsFirstDataGroup->setAutoFormat($xlsAutoFormat);
    $xls->easy_getSheetAt(0)->easy_addDataGroup($xlsFirstDataGroup);

    //Create the second group
    $xlsSecondDataGroup = new COM("EasyXLS.ExcelDataGroup");
    $xlsSecondDataGroup->setRange_2("A2:E10");
    $xlsSecondDataGroup->setGroupType($DATAGROUP_GROUP_BY_ROWS);
    $xlsSecondDataGroup->setCollapsed(False);
    
    //Create an instance of the object used to format the cells of the second group
    $xlsAutoFormat2 = new COM("EasyXLS.ExcelAutoFormat");
    $xlsAutoFormat2->InitAs($AUTOFORMAT_EASYXLS2);
    $xlsSecondDataGroup->setAutoFormat($xlsAutoFormat2);
    $xls->easy_getSheetAt(0)->easy_addDataGroup($xlsSecondDataGroup);


	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial17.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial17.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>