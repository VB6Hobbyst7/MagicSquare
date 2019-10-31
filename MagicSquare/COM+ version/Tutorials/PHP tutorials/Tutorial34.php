<?php
	/*==========================================================================
	 | Tutorial 34
	 |
	 | This tutorial shows how to read values from the active sheet
	 | of an excel file (the file generated in Tutorial 09).
	  ==========================================================================*/
	  
	header("Content-Type: text/html");
	
	echo "Tutorial 34<br>";
	echo "----------<br>";
	


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Read the file
	echo "Reading file: C:\\Samples\\Tutorial09.xls<br><br>";
	$rows = $xls->easy_ReadXLSActiveSheet_AsList("C:\\Samples\\Tutorial09.xls");
	
	//Confirm reading
	if ($xls->easy_getError() == "")
	{
		//Display the values
		for ($rowIndex=0; $rowIndex<$rows->size(); $rowIndex++)
		{
			$row = $rows->elementAt($rowIndex);
			for ($cellIndex=0; $cellIndex<$row->size(); $cellIndex++)
			{
				echo "At row ".($rowIndex + 1).", column ".($cellIndex + 1)." the value is '".$row->elementAt($cellIndex)."'<br>";
			}
		}
	}	
	else
		echo "Error reading file C:\Samples\Tutorial09.xls " . $xls->easy_getError();

	//Dispose memory
	$xls->Dispose();
?>

