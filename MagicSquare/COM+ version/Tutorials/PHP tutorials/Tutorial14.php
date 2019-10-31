<?php
	/*==========================================================================
	 | Tutorial 14
	 |
	 | This tutorial shows how to create conditional formatting ranges.
	  ==========================================================================*/
	
	//Include Files
	include("DataType.inc");
	include("ConditionalFormatting.inc");
	include("FontSettings.inc");
	include("Border.inc");
	include("Color.inc");

	header("Content-Type: text/html");
	
	echo "Tutorial 14<br>";
	echo "----------<br>";
	


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
		
	//Add a worksheet
	$xls->easy_addWorksheet_2("Sheet1");

	//Insert values
	$xlsTab = $xls->easy_getSheet("Sheet1");	
	$xlsTable = $xlsTab->easy_getExcelTable();

	for ($i=0; $i<6; $i++)
	{
		for ($j=0; $j<4; $j++)
		{
			if(($i<2)&&($j<2))
				$xlsTable->easy_getCell($i, $j)->setValue("12");
			else
				if(($j==2)&&($i<2))
					$xlsTable->easy_getCell($i, $j)->setValue("1000");
				else
					$xlsTable->easy_getCell($i, $j)->setValue("9");
			$xlsTable->easy_getCell($i, $j)->setDataType($DATATYPE_NUMERIC) ;
		}
	}

	//Set a conditional formatting
	$xlsTab->easy_addConditionalFormatting_5("A1:C3", $CONDITIONALFORMATTING_OPERATOR_BETWEEN, "=9", "=11", true, true, (int)$COLOR_RED);

	//Set a conditional formatting
	$xlsTab->easy_addConditionalFormatting_9("A6:C6", $CONDITIONALFORMATTING_OPERATOR_BETWEEN, "=COS(PI())+2", "", (int)$COLOR_BISQUE);
	$xlsTab->easy_getConditionalFormattingAt_2("A6:C6")->getConditionAt(0)->setConditionType($CONDITIONALFORMATTING_CONDITIONAL_FORMATTING_TYPE_EVALUATE_FORMULA);
	

	//Generate the file
	echo "Writing file: C:\Samples\Tutorial14.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial14.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>