<?php
	/*==========================================================================
	 | Tutorial 35
	 |
	 | This tutorial shows how to read values from a sheet
	 | of an excel file (For this example we use the file generated
	 | in Tutorial 09).
	  ==========================================================================*/
	
	header("Content-Type: text/html");
	
	echo "Tutorial 35<br>";
	echo "----------<br>";
		


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Read the file
	echo "Reading file: C:\\Samples\\Tutorial09.xls<br><br>";
	$rows = $xls->easy_ReadXLSSheet_AsList_3("C:\\Samples\\Tutorial09.xls", "First tab");

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
