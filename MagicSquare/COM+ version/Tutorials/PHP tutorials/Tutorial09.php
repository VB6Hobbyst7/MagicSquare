<?php
	/*==========================================================================
	 | Tutorial 09
	 |
	 | This tutorial shows how to create a Microsoft Excel file
	 | that has two worksheets. The first one is full with data
	 | and the cells are formatted. The column header has comments.
	 | The first worksheet has header & footer. The print options are
	 | set for the first worksheet.
	  ==========================================================================*/
	
	//Include files containing constants
	include("DataType.inc");
	include("Alignment.inc");
	include("Border.inc");
	include("Color.inc");
	include("Header.inc");
	include("Footer.inc");
	include("PageSetup.inc");

	header("Content-Type: text/html");
	
	echo "Tutorial 09<br>";
	echo "----------<br>";
	


	//Create an instance of the object that generates Excel files
	$xls = new COM("EasyXLS.ExcelDocument");
	
	//Create the worksheets 
	$xls->easy_addWorksheet_2("First tab");
	$xls->easy_addWorksheet_2("Second tab");

	//Lock the first tab
	$xls->easy_getSheetAt(0)->setSheetProtected(true);
	
	//Get the table of the first worksheet
	$xlsFirstTable = $xls->easy_getSheetAt(0)->easy_getExcelTable();

	//Create the style for the header
	$xlsStyleHeader = new COM("EasyXLS.ExcelStyle");
	$xlsStyleHeader->setFont("Verdana");
	$xlsStyleHeader->setFontSize(8);
	$xlsStyleHeader->setItalic(True);
	$xlsStyleHeader->setBold(True);
	$xlsStyleHeader->setForeground((int)$COLOR_YELLOW);
	$xlsStyleHeader->setBackground((int)$COLOR_BLACK);
	$xlsStyleHeader->setBorderColors ((int)$COLOR_GRAY, (int)$COLOR_GRAY, (int)$COLOR_GRAY, (int)$COLOR_GRAY);
	$xlsStyleHeader->setBorderStyles ($BORDER_BORDER_MEDIUM, $BORDER_BORDER_MEDIUM, $BORDER_BORDER_MEDIUM, $BORDER_BORDER_MEDIUM);
	$xlsStyleHeader->setHorizontalAlignment($ALIGNMENT_ALIGNMENT_CENTER);
	$xlsStyleHeader->setVerticalAlignment($ALIGNMENT_ALIGNMENT_BOTTOM);
	$xlsStyleHeader->setWrap(True);
	$xlsStyleHeader->setDataType($DATATYPE_STRING);
	
	//Add the cells for header
	for ($column=0; $column<5; $column++)
	{
		$xlsFirstTable->easy_getCell(0,$column)->setValue("Column " . ($column + 1));
		$xlsFirstTable->easy_getCell(0,$column)->setStyle($xlsStyleHeader);

		//Add comment
		$xlsFirstTable->easy_getCell(0, $column)->setComment_2("This is column no " . ($column + 1));
	}
	$xlsFirstTable->easy_getRowAt(0)->setHeight(30);
	
	//Create a style for cells
	$xlsStyleData = new COM("EasyXLS.ExcelStyle");
	$xlsStyleData->setHorizontalAlignment($ALIGNMENT_ALIGNMENT_LEFT);
	$xlsStyleData->setForeground((int)$COLOR_DARKGRAY);
	$xlsStyleData->setWrap(false);
	$xlsStyleData->setLocked(true);
	$xlsStyleData->setDataType($DATATYPE_STRING);
	
	//Add the cells for data
	for ($row=0; $row<100; $row++)
	{
		for ($column=0; $column<5; $column++)
		{
			$xlsFirstTable->easy_getCell($row+1,$column)->setValue("Data " . ($row + 1) . ", " . ($column + 1));
			$xlsFirstTable->easy_getCell($row+1,$column)->setStyle($xlsStyleData);
		}
	}
	
	//Set column widths
	$xlsFirstTable->setColumnWidth_2(0, 70);
	$xlsFirstTable->setColumnWidth_2(1, 100);
	$xlsFirstTable->setColumnWidth_2(2, 70);
	$xlsFirstTable->setColumnWidth_2(3, 100);
	$xlsFirstTable->setColumnWidth_2(4, 70);
	
	//Add headers for the first worksheet
	$xlsFirstTab = $xls->easy_getSheetAt(0);
	$xlsFirstTab->easy_getHeaderAt_2($HEADER_POSITION_CENTER)->InsertSingleUnderline();
	$xlsFirstTab->easy_getHeaderAt_2($HEADER_POSITION_CENTER)->InsertFile();
	$xlsFirstTab->easy_getHeaderAt_2($HEADER_POSITION_CENTER)->InsertValue(" - How to create header and footer");

	$xlsFirstTab->easy_getHeaderAt_2($HEADER_POSITION_RIGHT)->InsertDate();
	$xlsFirstTab->easy_getHeaderAt_2($HEADER_POSITION_RIGHT)->InsertValue(" ");
	$xlsFirstTab->easy_getHeaderAt_2($HEADER_POSITION_RIGHT)->InsertTime();

	//Add footer for the first worksheet
	$xlsFirstTab->easy_getFooterAt_2($FOOTER_POSITION_CENTER)->InsertPage();
	$xlsFirstTab->easy_getFooterAt_2($FOOTER_POSITION_CENTER)->InsertValue(" of ");
	$xlsFirstTab->easy_getFooterAt_2($FOOTER_POSITION_CENTER)->InsertPages();
	
	//Set Page Setup options
	$xlsPageSetup = $xlsFirstTab->easy_getPageSetup();
	$xlsPageSetup->easy_setPrintArea_3("A1:E101");
	$xlsPageSetup->easy_setRowsToRepeatAtTop_3("$1:$1");
	$xlsPageSetup->setCenterHorizontally(true);
	$xlsPageSetup->setOrientation($PAGESETUP_ORIENTATION_PORTRAIT);
	$xlsPageSetup->setPageOrder($PAGESETUP_PAGE_ORDER_DOWN_THEN_OVER);
	$xlsPageSetup->setPaperSize($PAGESETUP_PAPER_SIZE_A4);
	$xlsPageSetup->setPrintComments($PAGESETUP_COMMENTS_AT_END_OF_SHEET);
	$xlsPageSetup->setPrintGridlines(true);
	$xlsFirstTable->easy_insertPageBreakAtRow(21);
	$xlsFirstTable->easy_insertPageBreakAtRow(41);
	$xlsFirstTable->easy_insertPageBreakAtRow(61);
	$xlsFirstTable->easy_insertPageBreakAtRow(81);
	$xlsFirstTab->setPageBreakPreview(true);
	
	//Generate the file
	echo "Writing file: C:\Samples\Tutorial09.xls<br>";
	$xls->easy_WriteXLSFile("C:\Samples\Tutorial09.xls");
	
	//Confirm generation
	if ($xls->easy_getError() == "")
		echo "File successfully created.";
	else
		echo "Error encountered: " . $xls->easy_getError();
		
	//Dispose memory
	$xls->Dispose();
?>