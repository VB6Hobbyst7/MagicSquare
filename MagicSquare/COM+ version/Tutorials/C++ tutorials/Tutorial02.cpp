/* ----------------------------------------------------------------
 * Tutorial 02
 *
 * This tutorial shows how to generate an Excel document from a list of values. 
 * The cells are formatted using an user-defined format.
 * ----------------------------------------------------------------- */

#include "stdafx.h"
#include "EasyXLS.h"
#include <conio.h>
#import "C:\Program Files\Common Files\System\ado\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")

#define 	OddRowStripesStyleColor   0xfff0f7ef

int _tmain(int argc, _TCHAR* argv[])
{

	printf("Tutorial 02\n----------\n");

	HRESULT hr ;

	//Initialize COM
	hr = CoInitialize(0);
	
	// Use the SUCCEEDED macro and get a pointer to the interface
	if(SUCCEEDED(hr))
	{
		//Create a pointer to the interface that generates Excel files
		EasyXLS::IExcelDocumentPtr xls;
		hr = CoCreateInstance(__uuidof(EasyXLS::ExcelDocument),
		NULL,
		CLSCTX_ALL,
		__uuidof(EasyXLS::IExcelDocument),
		(void**) &xls) ;

		if(SUCCEEDED(hr))
		{
			// Connect to the database
			_ConnectionPtr objConn;
			objConn.CreateInstance(__uuidof(Connection));
			objConn->Open("driver={sql server};server=(local);Database=Northwind;UID=sa;PWD=;", (BSTR) NULL, (BSTR) NULL, -1);
									
			WCHAR* sQueryString = L"SELECT TOP 100 CAST(Month(ord.OrderDate) AS varchar)+'/' + CAST(Day(ord.OrderDate) AS varchar) + '/' + CAST(year(ord.OrderDate) AS varchar) AS 'Order Date', P.ProductName AS 'Product Name', O.UnitPrice AS Price, ' ' + cast(O.Quantity AS varchar) AS Quantity , O.UnitPrice * O. Quantity AS Value FROM Orders AS ord, [Order Details] AS O, Products AS P WHERE 	O.ProductID = P.ProductID AND O.OrderID = ord.OrderID";
			_variant_t sqlQueryString = sQueryString ;

			// Create the record set object
			_RecordsetPtr objRS = NULL;
			objRS.CreateInstance(__uuidof(Recordset));
			objRS->Open( sqlQueryString, _variant_t((IDispatch*)objConn,true), adOpenStatic, adLockOptimistic, adCmdText);
			
			// Create the list used to store the values	
			EasyXLS::IListPtr lstRows;
			CoCreateInstance(__uuidof(EasyXLS::List), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IList), (void**) &lstRows) ;
		
			// Add the header row to the list	
			EasyXLS::IListPtr lstHeaderRow;
			CoCreateInstance(__uuidof(EasyXLS::List), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IList), (void**) &lstHeaderRow) ;		
			lstHeaderRow->addElement("Order Date");
			lstHeaderRow->addElement("Product Name");
			lstHeaderRow->addElement("Price");
			lstHeaderRow->addElement("Quantity");
			lstHeaderRow->addElement("Value");
			lstRows->addElement(_variant_t((IDispatch*)lstHeaderRow,true));


			VARIANT index;
			index.vt=VT_I4;
			FieldPtr field;	
			
			//Add the values from the database to the list
			while (!(objRS->EndOfFile))
			{
				EasyXLS::IListPtr  RowList;
				CoCreateInstance(__uuidof(EasyXLS::List), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IList), (void**) &RowList) ;
				VARIANT value;					

				for (int nIndex = 0; nIndex < 5; nIndex++)
				{
					index.lVal = nIndex;
					objRS->Fields->get_Item(index, &field);
					field->get_Value (&value);
					RowList->addElement(&value);
				}			
				lstRows->addElement(_variant_t((IDispatch*)RowList,true));
						
				//Move to the next record
				objRS->MoveNext();
			}

			// Create an instance of the object used to format the cells
			EasyXLS::IExcelAutoFormatPtr xlsAutoFormat;
			CoCreateInstance(__uuidof(EasyXLS::ExcelAutoFormat), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelAutoFormat), (void**) &xlsAutoFormat) ;
			// Set the style of the header
			EasyXLS::IExcelStylePtr xlsHeaderStyle;
			CoCreateInstance(__uuidof(EasyXLS::ExcelStyle), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelStyle), (void**) &xlsHeaderStyle);
			xlsHeaderStyle->setBackground(COLOR_LIGHTGREEN);
			xlsHeaderStyle->setFontSize(12);
			xlsAutoFormat->setHeaderRowStyle(xlsHeaderStyle);

			// Set the style of the cells
			EasyXLS::IExcelStylePtr xlsEvenRowStripesStyle;
			CoCreateInstance(__uuidof(EasyXLS::ExcelStyle), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelStyle), (void**) &xlsEvenRowStripesStyle);
			xlsEvenRowStripesStyle->setBackground(COLOR_FLORALWHITE);
			xlsEvenRowStripesStyle->setFormat("$0.00");
			xlsEvenRowStripesStyle->setHorizontalAlignment(ALIGNMENT_ALIGNMENT_LEFT);
			xlsAutoFormat->setEvenRowStripesStyle(xlsEvenRowStripesStyle)	;
			EasyXLS::IExcelStylePtr xlsOddRowStripesStyle;
			CoCreateInstance(__uuidof(EasyXLS::ExcelStyle), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelStyle), (void**) &xlsOddRowStripesStyle);
			xlsOddRowStripesStyle->setBackground(OddRowStripesStyleColor);
			xlsOddRowStripesStyle->setFormat("$0.00");
			xlsOddRowStripesStyle->setHorizontalAlignment (ALIGNMENT_ALIGNMENT_LEFT);
			xlsAutoFormat->setOddRowStripesStyle(xlsOddRowStripesStyle);
			EasyXLS::IExcelStylePtr xlsLeftColumnStyle;
			CoCreateInstance(__uuidof(EasyXLS::ExcelStyle), NULL, CLSCTX_ALL, __uuidof(EasyXLS::IExcelStyle), (void**) &xlsLeftColumnStyle);
			xlsLeftColumnStyle->setBackground(COLOR_FLORALWHITE);
			xlsLeftColumnStyle->setFormat("mm/dd/yyyy");
			xlsLeftColumnStyle->setHorizontalAlignment(ALIGNMENT_ALIGNMENT_LEFT);
			xlsAutoFormat->setLeftColumnStyle(xlsLeftColumnStyle)	;

			//Generate the file
			printf("Writing file C:\\Samples\\Tutorial02.xls.");
			hr = xls->easy_WriteXLSFile_FromList_2("C:\\Samples\\Tutorial02.xls", _variant_t((IDispatch*)lstRows,true),  _variant_t((IDispatch*)xlsAutoFormat,true), "Sheet1");				
			
			//Confirm generation
			_bstr_t sError = xls->easy_getError();
			if (strcmp(sError, "") == 0)
			{
				printf("\nFile successfully created. Press Enter to Exit...");
			}
			else
			{
				printf("\nError encountered: %s", (LPCSTR)sError); 
			}
						
			// Close the Recordset object
			objRS->Close();
		
			// Close the Connection object
			objConn->Close();

			//Dispose memory
			xls->Dispose();
		}
   		else
		{
			printf("Object is not available!");
		}
	}
	else{
		printf("COM can't be initialized!");
	}

	// Uninitialize COM
	CoUninitialize();

	_getch();
	return 0;

}