<?php
	/*==========================================================================

	 | Tutorial 33
	 |
	 | This tutorial shows how to set the properties of the document.

==========================================================================*/

	//Include files containing constants
	include("FileProperty.inc");

	header("Content-Type: text/html");

	echo "Tutorial 33<br>";
	echo "----------<br>";

//Create an instance of the object that generates Excel files
$xls = new COM("EasyXLS.ExcelDocument");

//Add the worksheet
$xls->easy_addWorksheet_2("Sheet1");

//Set the 'Subject' property
$xls->getSummaryInformation()->setSubject("This is the subject");

//Set the 'Manager' property
$xls->getDocumentSummaryInformation()->setManager("This is the manager");

//Set a custom property
$xls->getDocumentSummaryInformation()->setCustomProperty("PropertyName", $VT_NUMBER, "4");

//Generate the file
echo "Writing file: C:\Samples\Tutorial33.xls<br>";
$xls->easy_WriteXLSFile("C:\Samples\Tutorial33.xls");

//Confirm generation
if ($xls->easy_getError() == "")
		        echo "File successfully created.";
		else
		        echo "Error encountered: " .$xls->easy_getError();

//Dispose memory
	$xls->Dispose();	
?>
