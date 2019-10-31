/* ----------------------------------------------------------------
 | Tutorial 34                                                    
 |                                                                
 | This tutorial shows how to read values from the active sheet   
 | of an excel file (the file generated in Tutorial 09).          
 -----------------------------------------------------------------*/

#include "stdafx.h"

using namespace System;
using namespace System::IO;
using namespace System::Data;
using namespace EasyXLS;
using namespace System::Text;

int main()
{

	Console::WriteLine("Tutorial 34\n----------\n");

	//Create an instance of the object that generates Excel files, having 2 sheets	
	ExcelDocument ^xls = gcnew ExcelDocument();

	//Read the file
	Console::WriteLine("Reading file C:\\Samples\\Tutorial09.xls.\n");
	DataSet ^ds = xls->easy_ReadXLSActiveSheet_AsDataSet("C:\\Samples\\Tutorial09.xls");

	//Confirm generation
	String ^sError = xls->easy_getError();
	if (sError->Equals(""))
	{
		//Display the values
		DataTable ^dt = ds->Tables[0];
		StringBuilder ^str;
		for (int row=0; row < dt->Rows->Count; row++)
		{
			for (int column=0; column < dt->Columns->Count; column++)
			{
				str = gcnew StringBuilder();
				str->Append(String::Concat("At row ", (row + 1).ToString(), ", column ", (column + 1).ToString()));
				str->Append(String::Concat(" the value is '", Convert::ToString(dt->Rows[row]->ItemArray[column]), "'"));
				Console::WriteLine(str);
			}
		}
	}
	else
		Console::Write(String::Concat("\nError reading file C:\\Samples\\Tutorial09.xls \n", sError));

	//Dispose memory
    delete xls;
	
	Console::Write("\nPress Enter to Exit...");
	Console::ReadLine();

	return 0;
}