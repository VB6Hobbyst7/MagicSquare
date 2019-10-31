/* ----------------------------------------------------------------
 * Tutorial 08
 * 
 * This tutorial shows how to create a Microsoft Excel file
 * that has two worksheets. The first one is full with data
 * and the cells are formatted. The column header has comments.
 * The first worksheet has header & footer.
 * ----------------------------------------------------------------- */

using System;
using System.Drawing;
using EasyXLS;
using EasyXLS.Constants;

public class Tutorial08
{
	
	[STAThread]
	static void Main() 
	{
		Console.WriteLine("Tutorial 08\n-----------\n");

		//Create an instance of the object that generates Excel files, having 2 sheets
		ExcelDocument xls = new ExcelDocument(2);
	    
		//Set the sheet names
		xls.easy_getSheetAt(0).setSheetName("First tab");
		xls.easy_getSheetAt(1).setSheetName("Second tab");

		//Lock the first tab
		xls.easy_getSheetAt(0).setSheetProtected(true);

		//Get the table of the first worksheet
		ExcelTable xlsFirstTable = ((ExcelWorksheet)xls.easy_getSheetAt(0)).easy_getExcelTable();

		//Create the style for the header
		ExcelStyle xlsStyleHeader = new ExcelStyle("Verdana", 8, true, true, Color.Yellow);		
		xlsStyleHeader.setBackground(Color.Black);
		xlsStyleHeader.setBorderColors(Color.Gray, Color.Gray, Color.Gray, Color.Gray);
		xlsStyleHeader.setBorderStyles(Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM, Border.BORDER_MEDIUM);	
		xlsStyleHeader.setHorizontalAlignment(Alignment.ALIGNMENT_CENTER);
		xlsStyleHeader.setVerticalAlignment(Alignment.ALIGNMENT_BOTTOM);
		xlsStyleHeader.setWrap(true);
		xlsStyleHeader.setDataType(DataType.STRING);

		//Add cells for header
		for (int column=0; column<5; column++)
		{
			xlsFirstTable.easy_getCell(0, column).setValue("Column " + (column + 1)); 
			xlsFirstTable.easy_getCell(0, column).setStyle(xlsStyleHeader);

			//Add comment
			xlsFirstTable.easy_getCell(0, column).setComment("This is column no " + (column + 1));
		}
		xlsFirstTable.easy_getRowAt(0).setHeight(30);

		//Add cells for data
		for (int row=0; row<100; row++)
		{
			for (int column=0; column<5; column++)
			{
				xlsFirstTable.easy_getCell(row+1, column).setValue("Data " + (row + 1) + ", " + (column + 1));
			}
		}

		//Create a style for cells
		ExcelStyle xlsStyleData = new ExcelStyle();
		xlsStyleData.setHorizontalAlignment(Alignment.ALIGNMENT_LEFT);
		xlsStyleData.setForeground(Color.DarkGray);
		xlsStyleData.setWrap(false);
		xlsStyleData.setDataType(DataType.STRING);
		xlsStyleData.setLocked(true);
		xlsFirstTable.easy_setRangeStyle("A2:E101", xlsStyleData);

		//Set column widths
		xlsFirstTable.setColumnWidth(0, 70);
		xlsFirstTable.setColumnWidth(1, 100);
		xlsFirstTable.setColumnWidth(2, 70);
		xlsFirstTable.setColumnWidth(3, 100);
		xlsFirstTable.setColumnWidth(4, 70);

		//Add headers for the first worksheet
		ExcelWorksheet xlsFirstTab = (ExcelWorksheet)xls.easy_getSheetAt(0);
		xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertSingleUnderline();
		xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertFile();
		xlsFirstTab.easy_getHeaderAt(Header.POSITION_CENTER).InsertValue(" - How to create header and footer");

		xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertDate();
		xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertValue(" ");
		xlsFirstTab.easy_getHeaderAt(Header.POSITION_RIGHT).InsertTime();

		//Add footer for the first worksheet
		xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertPage();
		xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertValue(" of ");
		xlsFirstTab.easy_getFooterAt(Footer.POSITION_CENTER).InsertPages();

		//Generate the file
		Console.WriteLine("Writing file C:\\Samples\\Tutorial08.xls.");
		xls.easy_WriteXLSFile("C:\\Samples\\Tutorial08.xls");

		//Confirm generation
		String sError = xls.easy_getError();
		if (sError.Equals(""))
			Console.Write("\nFile successfully created. Press Enter to Exit...");
		else
			Console.Write("\nError encountered: " + sError + "\nPress Enter to Exit...");
		
		//Dispose memory
		xls.Dispose();

		Console.ReadLine();
	}
}

