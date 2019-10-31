<?php
	/*==========================================================================
	 | Tutorial 19
	 |
	 | This tutorial shows how to create a Microsoft Excel file
	 | that has two worksheets. The first one is full with data
	 | and the first cell of the second row contains Rich Text Format.
	  ==========================================================================*/
	
	//Include files containing constants
	include("DataType.inc");

	header("Content-Type: text/html");
	
	echo "Tutorial 19<br>";
	echo "----------<br>";
	


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Create the worksheets
	$xls->easy_addWorksheet_2("First tab");
	$xls->easy_addWorksheet_2("Second tab");

	//Get the table of the first worksheet
	$xlsFirstTable = $xls->easy_getSheetAt(0)->easy_getExcelTable();

    //Create the string used to set the RTF in cell
    $sFormattedValue = "This is <b>bold</b>.";
    $sFormattedValue = $sFormattedValue . "\nThis is <i>italic</i>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <u>underline</u>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <underline double>double underline</underline double>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <font color=red>red</font>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <font color=rgb(255,0,0)>red</font> too.";
    $sFormattedValue = $sFormattedValue ."\nThis is <font face=\"Arial Black\">Arial Black</font>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <font size=15pt>size 15</font>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <s>strikethrough</s>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <sup>superscript</sup>.";
    $sFormattedValue = $sFormattedValue ."\nThis is <sub>subscript</sub>.";
    $sFormattedValue = $sFormattedValue ."\n<b>This</b> <i>is</i> <font color=red face=\"Arial Black\" size=15pt><underline double>formatted</underline double></font> <s>text</s>.";


    //Set the formatted value
    $xlsFirstTable->easy_getCell(1, 0)->setHTMLValue ($sFormattedValue);
    $xlsFirstTable->easy_getCell(1, 0)->setDataType ($DATATYPE_STRING);
    $xlsFirstTable->easy_getCell(1, 0)->setWrap (True);
    $xlsFirstTable->easy_getRowAt(1)->setHeight (250);
    $xlsFirstTable->easy_getColumnAt(0)->setWidth (250);



	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial19.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial19.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>