<?php
	/*==========================================================================
	 | Tutorial 25
	 |
	 | This tutorial shows how to create a pivot table.
	  ==========================================================================*/
	
	//Include files containing constants
	include("DataType.inc");
	include("PivotTable.inc");

	header("Content-Type: text/html");

	echo "Tutorial 25<br>";
	echo "----------<br>";
	

	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Create the worksheets
	$xls->easy_addWorksheet_2("First tab");
	$xls->easy_addWorksheet_2("Second tab");

	//Get the table of the first worksheet
	$xlsFirstTable = $xls->easy_getSheetAt(0)->easy_getExcelTable();

	//Add the cells for header
	$xlsFirstTable->easy_getCell(0,0)->setValue("Sale agent");
	$xlsFirstTable->easy_getCell(0,0)->setDataType($DATATYPE_STRING);
	$xlsFirstTable->easy_getCell(0,1)->setValue("Sale country");
	$xlsFirstTable->easy_getCell(0,1)->setDataType($DATATYPE_STRING);
	$xlsFirstTable->easy_getCell(0,2)->setValue("Month");
	$xlsFirstTable->easy_getCell(0,2)->setDataType($DATATYPE_STRING);
	$xlsFirstTable->easy_getCell(0,3)->setValue("Year");
	$xlsFirstTable->easy_getCell(0,3)->setDataType($DATATYPE_STRING);
	$xlsFirstTable->easy_getCell(0,4)->setValue("Sale amount");
	$xlsFirstTable->easy_getCell(0,4)->setDataType($DATATYPE_STRING);
		
	$xlsFirstTable->easy_getRowAt(0)->setBold(true);

	//Populate the source for pivot table
	$xlsFirstTable->easy_getCell(1,0)->setValue("John Down");
	$xlsFirstTable->easy_getCell(1,1)->setValue("USA");
	$xlsFirstTable->easy_getCell(1,2)->setValue("June");
	$xlsFirstTable->easy_getCell(1,3)->setValue("2010");
	$xlsFirstTable->easy_getCell(1,4)->setValue("550");
		
	$xlsFirstTable->easy_getCell(2,0)->setValue("Scott Valey");
	$xlsFirstTable->easy_getCell(2,1)->setValue("United Kingdom");
	$xlsFirstTable->easy_getCell(2,2)->setValue("June");
	$xlsFirstTable->easy_getCell(2,3)->setValue("2010");
	$xlsFirstTable->easy_getCell(2,4)->setValue("2300");
		
	$xlsFirstTable->easy_getCell(3,0)->setValue("John Down");
	$xlsFirstTable->easy_getCell(3,1)->setValue("USA");
	$xlsFirstTable->easy_getCell(3,2)->setValue("July");
	$xlsFirstTable->easy_getCell(3,3)->setValue("2010");
	$xlsFirstTable->easy_getCell(3,4)->setValue("3100");
		
	$xlsFirstTable->easy_getCell(4,0)->setValue("John Down");
	$xlsFirstTable->easy_getCell(4,1)->setValue("USA");
	$xlsFirstTable->easy_getCell(4,2)->setValue("June");
	$xlsFirstTable->easy_getCell(4,3)->setValue("2011");
	$xlsFirstTable->easy_getCell(4,4)->setValue("1050");
			
	$xlsFirstTable->easy_getCell(5,0)->setValue("John Down");
	$xlsFirstTable->easy_getCell(5,1)->setValue("USA");
	$xlsFirstTable->easy_getCell(5,2)->setValue("July");
	$xlsFirstTable->easy_getCell(5,3)->setValue("2011");
	$xlsFirstTable->easy_getCell(5,4)->setValue("2400");
		
	$xlsFirstTable->easy_getCell(6,0)->setValue("Steve Marlowe");
	$xlsFirstTable->easy_getCell(6,1)->setValue("France");
	$xlsFirstTable->easy_getCell(6,2)->setValue("June");
	$xlsFirstTable->easy_getCell(6,3)->setValue("2011");
	$xlsFirstTable->easy_getCell(6,4)->setValue("1200");
		
	$xlsFirstTable->easy_getCell(7,0)->setValue("Scott Valey");
	$xlsFirstTable->easy_getCell(7,1)->setValue("United Kingdom");
	$xlsFirstTable->easy_getCell(7,2)->setValue("June");
	$xlsFirstTable->easy_getCell(7,3)->setValue("2011");
	$xlsFirstTable->easy_getCell(7,4)->setValue("700");
		
	$xlsFirstTable->easy_getCell(8,0)->setValue("Scott Valey");
	$xlsFirstTable->easy_getCell(8,1)->setValue("United Kingdom");
	$xlsFirstTable->easy_getCell(8,2)->setValue("July");
	$xlsFirstTable->easy_getCell(8,3)->setValue("2011");
	$xlsFirstTable->easy_getCell(8,4)->setValue("360");
	
	//Create pivot table
	$xlsPivotTable = new COM("EasyXLS.PivotTables.ExcelPivotTable");
		
	$xlsPivotTable->setName("Sales");
	$xlsPivotTable->setSourceRange("First tab!\$A\$1:\$E\$9", $xls);
	$xlsPivotTable->setLocation_2("A3:G15");
	$xlsPivotTable->addFieldToRowLabels("Sale agent");
	$xlsPivotTable->addFieldToColumnLabels("Year");
	$xlsPivotTable->addFieldToValues("Sale amount","Sale amount per year",$PIVOTTABLE_SUBTOTAL_SUM);  
	$xlsPivotTable->addFieldToReportFilter("Sale country");
	$xlsPivotTable->setOutlineForm(); 
	$xlsPivotTable->setStyle($PIVOTTABLE_PIVOT_STYLE_DARK_11);
			
	//Add the pivot table
	$xlsWorksheet = $xls->easy_getSheet("Second tab");
	$xlsWorksheet->easy_addPivotTable($xlsPivotTable);
			
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial25.xlsx<br>";
	$xls->easy_WriteXLSXFile("C:\Samples\Tutorial25.xlsx");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>