#pragma once


namespace CPlusPlusDotNetApplication
{
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	
	using namespace EasyXLS;
	using namespace EasyXLS::Charts;
	using namespace EasyXLS::Constants;

	/// <summary> 
	/// Summary for Form1
	///
	/// WARNING: If you change the name of this class, you will need to change the 
	///          'Resource File Name' property for the managed resource compiler tool 
	///          associated with all .resx files this class depends on.  Otherwise,
	///          the designers will not be able to interact properly with localized
	///          resources associated with this form.
	/// </summary>
	public __gc class Form1 : public System::Windows::Forms::Form
	{	
	public:
		Form1(void)
		{
			InitializeComponent();
		}
  
	protected:
		void Dispose(Boolean disposing)
		{
			if (disposing && components)
			{
				components->Dispose();
			}
			__super::Dispose(disposing);
		}
	private: System::Data::DataSet *  dsSource;
	private: System::Windows::Forms::DataGrid *  dgTimeSheetReport;

	private: System::Data::DataColumn *  dataColumn1;
	private: System::Data::DataColumn *  dataColumn2;
	private: System::Data::DataColumn *  dataColumn3;
	private: System::Data::DataColumn *  dataColumn4;
	private: System::Data::DataColumn *  dataColumn5;
	private: System::Data::DataColumn *  dataColumn6;
	private: System::Data::DataColumn *  dataColumn7;
	private: System::Data::DataColumn *  dataColumn8;
	private: System::Data::DataColumn *  dataColumn9;
	private: System::Data::DataTable *  dtTable;
	private: System::Windows::Forms::Button *  btnExporttoExcel;
	private: System::Windows::Forms::PictureBox *  imgEasyXLSlogo;
	private: System::Windows::Forms::LinkLabel *  hlkEasyXLS;
	private: System::Windows::Forms::Label *  label1;
	private: System::Windows::Forms::Label *  label2;
	private: System::Windows::Forms::Label *  label3;
	private: System::Windows::Forms::Label *  label4;
	private: System::Windows::Forms::CheckBox *  chkTask;
	private: System::Windows::Forms::CheckBox *  chkEstimated;
	private: System::Windows::Forms::CheckBox *  chkRegular;
	private: System::Windows::Forms::CheckBox *  chkOTHours;
	private: System::Windows::Forms::CheckBox *  chkNBHours;
	private: System::Windows::Forms::Label *  label5;



	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container * components;

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			System::Resources::ResourceManager *  resources = new System::Resources::ResourceManager(__typeof(CPlusPlusDotNetApplication::Form1));
			this->dsSource = new System::Data::DataSet();
			this->dtTable = new System::Data::DataTable();
			this->dataColumn1 = new System::Data::DataColumn();
			this->dataColumn2 = new System::Data::DataColumn();
			this->dataColumn3 = new System::Data::DataColumn();
			this->dataColumn4 = new System::Data::DataColumn();
			this->dataColumn5 = new System::Data::DataColumn();
			this->dataColumn6 = new System::Data::DataColumn();
			this->dataColumn7 = new System::Data::DataColumn();
			this->dataColumn8 = new System::Data::DataColumn();
			this->dataColumn9 = new System::Data::DataColumn();
			this->dgTimeSheetReport = new System::Windows::Forms::DataGrid();
			this->btnExporttoExcel = new System::Windows::Forms::Button();
			this->imgEasyXLSlogo = new System::Windows::Forms::PictureBox();
			this->hlkEasyXLS = new System::Windows::Forms::LinkLabel();
			this->label1 = new System::Windows::Forms::Label();
			this->label2 = new System::Windows::Forms::Label();
			this->label3 = new System::Windows::Forms::Label();
			this->label4 = new System::Windows::Forms::Label();
			this->chkTask = new System::Windows::Forms::CheckBox();
			this->chkEstimated = new System::Windows::Forms::CheckBox();
			this->chkRegular = new System::Windows::Forms::CheckBox();
			this->chkOTHours = new System::Windows::Forms::CheckBox();
			this->chkNBHours = new System::Windows::Forms::CheckBox();
			this->label5 = new System::Windows::Forms::Label();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->dsSource))->BeginInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->dtTable))->BeginInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->dgTimeSheetReport))->BeginInit();
			this->SuspendLayout();
			// 
			// dsSource
			// 
			this->dsSource->DataSetName = S"NewDataSet";
			this->dsSource->Locale = new System::Globalization::CultureInfo(S"en-US");
			System::Data::DataTable* __mcTemp__1[] = new System::Data::DataTable*[1];
			__mcTemp__1[0] = this->dtTable;
			this->dsSource->Tables->AddRange(__mcTemp__1);
			// 
			// dtTable
			// 
			System::Data::DataColumn* __mcTemp__2[] = new System::Data::DataColumn*[9];
			__mcTemp__2[0] = this->dataColumn1;
			__mcTemp__2[1] = this->dataColumn2;
			__mcTemp__2[2] = this->dataColumn3;
			__mcTemp__2[3] = this->dataColumn4;
			__mcTemp__2[4] = this->dataColumn5;
			__mcTemp__2[5] = this->dataColumn6;
			__mcTemp__2[6] = this->dataColumn7;
			__mcTemp__2[7] = this->dataColumn8;
			__mcTemp__2[8] = this->dataColumn9;
			this->dtTable->Columns->AddRange(__mcTemp__2);
			this->dtTable->TableName = S"dtTable";
			// 
			// dataColumn1
			// 
			this->dataColumn1->ColumnName = S"Project";
			// 
			// dataColumn2
			// 
			this->dataColumn2->ColumnName = S"Resource";
			// 
			// dataColumn3
			// 
			this->dataColumn3->ColumnName = S"Role";
			// 
			// dataColumn4
			// 
			this->dataColumn4->ColumnName = S"Task";
			// 
			// dataColumn5
			// 
			this->dataColumn5->ColumnName = S"Estimated";
			this->dataColumn5->DataType = __typeof(System::Int32);
			// 
			// dataColumn6
			// 
			this->dataColumn6->ColumnName = S"Regular";
			this->dataColumn6->DataType = __typeof(System::Int32);
			// 
			// dataColumn7
			// 
			this->dataColumn7->ColumnName = S"OT Hours";
			this->dataColumn7->DataType = __typeof(System::Int32);
			// 
			// dataColumn8
			// 
			this->dataColumn8->ColumnName = S"NB Hours";
			this->dataColumn8->DataType = __typeof(System::Int32);
			// 
			// dataColumn9
			// 
			this->dataColumn9->ColumnName = S"Approval Status";
			// 
			// dgTimeSheetReport
			// 
			this->dgTimeSheetReport->AllowSorting = false;
			this->dgTimeSheetReport->DataMember = S"dtTable";
			this->dgTimeSheetReport->DataSource = this->dsSource;
			this->dgTimeSheetReport->HeaderFont = new System::Drawing::Font(S"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point, (System::Byte)0);
			this->dgTimeSheetReport->HeaderForeColor = System::Drawing::SystemColors::ControlText;
			this->dgTimeSheetReport->Location = System::Drawing::Point(0, 156);
			this->dgTimeSheetReport->Name = S"dgTimeSheetReport";
			this->dgTimeSheetReport->PreferredColumnWidth = 90;
			this->dgTimeSheetReport->ReadOnly = true;
			this->dgTimeSheetReport->Size = System::Drawing::Size(892, 225);
			this->dgTimeSheetReport->TabIndex = 0;
			// 
			// btnExporttoExcel
			// 
			this->btnExporttoExcel->Font = new System::Drawing::Font(S"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point, (System::Byte)0);
			this->btnExporttoExcel->Location = System::Drawing::Point(24, 552);
			this->btnExporttoExcel->Name = S"btnExporttoExcel";
			this->btnExporttoExcel->Size = System::Drawing::Size(96, 24);
			this->btnExporttoExcel->TabIndex = 1;
			this->btnExporttoExcel->Text = S"Export xls file";
			this->btnExporttoExcel->Click += new System::EventHandler(this, btnExporttoExcel_Click);
			// 
			// imgEasyXLSlogo
			// 
			this->imgEasyXLSlogo->Image = (__try_cast<System::Drawing::Image *  >(resources->GetObject(S"imgEasyXLSlogo.Image")));
			this->imgEasyXLSlogo->Location = System::Drawing::Point(16, 8);
			this->imgEasyXLSlogo->Name = S"imgEasyXLSlogo";
			this->imgEasyXLSlogo->Size = System::Drawing::Size(154, 51);
			this->imgEasyXLSlogo->SizeMode = System::Windows::Forms::PictureBoxSizeMode::AutoSize;
			this->imgEasyXLSlogo->TabIndex = 2;
			this->imgEasyXLSlogo->TabStop = false;
			this->imgEasyXLSlogo->Tag = S"EasyXLSlogo.jpg";
			// 
			// hlkEasyXLS
			// 
			this->hlkEasyXLS->Font = new System::Drawing::Font(S"Microsoft Sans Serif", 10, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point, (System::Byte)0);
			this->hlkEasyXLS->Location = System::Drawing::Point(16, 96);
			this->hlkEasyXLS->Name = S"hlkEasyXLS";
			this->hlkEasyXLS->Size = System::Drawing::Size(168, 23);
			this->hlkEasyXLS->TabIndex = 3;
			this->hlkEasyXLS->TabStop = true;
			this->hlkEasyXLS->Text = S"http://www.easyxls.com";
			this->hlkEasyXLS->LinkClicked += new System::Windows::Forms::LinkLabelLinkClickedEventHandler(this, hlkEasyXLS_LinkClicked);
			// 
			// label1
			// 
			this->label1->Location = System::Drawing::Point(224, 16);
			this->label1->Name = S"label1";
			this->label1->Size = System::Drawing::Size(480, 104);
			this->label1->TabIndex = 4;
			this->label1->Click += new System::EventHandler(this, label1_Click);
			// 
			// label2
			// 
			this->label2->Font = new System::Drawing::Font(S"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Italic, System::Drawing::GraphicsUnit::Point, (System::Byte)0);
			this->label2->ForeColor = System::Drawing::SystemColors::ControlDarkDark;
			this->label2->Location = System::Drawing::Point(16, 72);
			this->label2->Name = S"label2";
			this->label2->TabIndex = 5;
			this->label2->Text = S"* sample image";
			// 
			// label3
			// 
			this->label3->Font = new System::Drawing::Font(S"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Italic, System::Drawing::GraphicsUnit::Point, (System::Byte)0);
			this->label3->ForeColor = System::Drawing::SystemColors::ControlDarkDark;
			this->label3->Location = System::Drawing::Point(16, 117);
			this->label3->Name = S"label3";
			this->label3->TabIndex = 6;
			this->label3->Text = S"* sample hyperlink";
			// 
			// label4
			// 
			this->label4->Font = new System::Drawing::Font(S"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Italic, System::Drawing::GraphicsUnit::Point, (System::Byte)0);
			this->label4->ForeColor = System::Drawing::SystemColors::ControlDarkDark;
			this->label4->Location = System::Drawing::Point(24, 384);
			this->label4->Name = S"label4";
			this->label4->Size = System::Drawing::Size(416, 23);
			this->label4->TabIndex = 7;
			this->label4->Text = S"* sample data set source; totals are computed using formulas";
			// 
			// chkTask
			// 
			this->chkTask->Checked = true;
			this->chkTask->CheckState = System::Windows::Forms::CheckState::Checked;
			this->chkTask->Enabled = false;
			this->chkTask->Location = System::Drawing::Point(48, 436);
			this->chkTask->Name = S"chkTask";
			this->chkTask->TabIndex = 8;
			this->chkTask->Text = S"Task";
			// 
			// chkEstimated
			// 
			this->chkEstimated->Checked = true;
			this->chkEstimated->CheckState = System::Windows::Forms::CheckState::Checked;
			this->chkEstimated->Location = System::Drawing::Point(48, 456);
			this->chkEstimated->Name = S"chkEstimated";
			this->chkEstimated->TabIndex = 9;
			this->chkEstimated->Text = S"Estimated";
			// 
			// chkRegular
			// 
			this->chkRegular->Checked = true;
			this->chkRegular->CheckState = System::Windows::Forms::CheckState::Checked;
			this->chkRegular->Location = System::Drawing::Point(48, 476);
			this->chkRegular->Name = S"chkRegular";
			this->chkRegular->TabIndex = 10;
			this->chkRegular->Text = S"Regular";
			// 
			// chkOTHours
			// 
			this->chkOTHours->Checked = true;
			this->chkOTHours->CheckState = System::Windows::Forms::CheckState::Checked;
			this->chkOTHours->Location = System::Drawing::Point(48, 496);
			this->chkOTHours->Name = S"chkOTHours";
			this->chkOTHours->TabIndex = 11;
			this->chkOTHours->Text = S"OT Hours";
			// 
			// chkNBHours
			// 
			this->chkNBHours->Checked = true;
			this->chkNBHours->CheckState = System::Windows::Forms::CheckState::Checked;
			this->chkNBHours->Location = System::Drawing::Point(48, 516);
			this->chkNBHours->Name = S"chkNBHours";
			this->chkNBHours->TabIndex = 12;
			this->chkNBHours->Text = S"NB Hours";
			// 
			// label5
			// 
			this->label5->Font = new System::Drawing::Font(S"Microsoft Sans Serif", 8.25F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point, (System::Byte)0);
			this->label5->Location = System::Drawing::Point(24, 416);
			this->label5->Name = S"label5";
			this->label5->Size = System::Drawing::Size(280, 23);
			this->label5->TabIndex = 13;
			this->label5->Text = S" Generate chart with the following columns:";
			// 
			// Form1
			// 
			this->AutoScaleBaseSize = System::Drawing::Size(5, 13);
			this->ClientSize = System::Drawing::Size(892, 597);
			this->Controls->Add(this->label5);
			this->Controls->Add(this->chkNBHours);
			this->Controls->Add(this->chkOTHours);
			this->Controls->Add(this->chkRegular);
			this->Controls->Add(this->chkEstimated);
			this->Controls->Add(this->chkTask);
			this->Controls->Add(this->label4);
			this->Controls->Add(this->label3);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->hlkEasyXLS);
			this->Controls->Add(this->imgEasyXLSlogo);
			this->Controls->Add(this->btnExporttoExcel);
			this->Controls->Add(this->dgTimeSheetReport);
			this->Name = S"Form1";
			this->Text = S"Form1";
			this->Load += new System::EventHandler(this, Form1_Load);
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->dsSource))->EndInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->dtTable))->EndInit();
			(__try_cast<System::ComponentModel::ISupportInitialize *  >(this->dgTimeSheetReport))->EndInit();
			this->ResumeLayout(false);

		}	
	private: System::Void Form1_Load(System::Object *  sender, System::EventArgs *  e)
			 {
				 // Populating the grid
                 Object* Row1[] = {S"EasyXLS", S"Jim Bean", S"Programmer", S"Build Charts", S"800", S"240", S"40", S"0", S"To be Approved"};
                 dtTable->Rows->Add(Row1);
                 Object* Row2[] = {S"EasyXLS", S"Jack White", S"Programmer", S"Build Worksheets", S"1000", S"160", S"0", S"0", S"To be Approved"};
                 dtTable->Rows->Add(Row2);
				 Object* Row3[] = {S"EasyXLS", S"Christina Brown", S"Programmer", S"Build Hyperlinks", S"750", S"256", S"2", S"0", S"To be Approved"};
                 dtTable->Rows->Add(Row3);
                 Object* Row4[]  = {S"EasyXLS", S"Walt Whitman", S"Programmer", S"Create Tutorials", S"600", S"114", S"10", S"0", S"To be Approved"};
                 dtTable->Rows->Add(Row4);
				 Object* Row5[] = {S"EasyXLS", S"Adam Wilson", S"Tester", S"Test Charts", S"120", S"8", S"0", S"0", S"To be Approved"};
                 dtTable->Rows->Add(Row5);
				 Object* Row6[] = {S"EasyXLS", S"Will Crane", S"Tester", S"Test Hyperlinks", S"100", S"10", S"2", S"0", S"To be Approved"};
                 dtTable->Rows->Add(Row6);
				 Object* Row7[] = {S"EasyXLS", S"George Brown", S"Artist", S"Design", S"300", S"150", S"2", S"0", S"To be Approved"};
                 dtTable->Rows->Add(Row7);
				 Object* Row8[] = {S"MS Excel", S"Christian Wurm", S"Programmer", S"Database Design", S"120", S"35", S"3", S"0", S"To be Approved"};
                 dtTable->Rows->Add(Row8);
				 Object* Row9[] = {S"MS Excel", S"Adrian Fisher", S"Tester", S"Speed", S"240", S"48", S"0", S"8", S"To be Approved"};
                 dtTable->Rows->Add(Row9);

				 Object* footerRow[] = {S"Totals:", S"", S"", S"", S"", S"", S"", S"", S""};

                 // Computing the totals
				 int nTotal = 0;
				 for (int nColumnIndex = 4; nColumnIndex < 8; nColumnIndex++)
				 {    
                     nTotal = 0;
					 for (int nRowIndex = 0; nRowIndex < dtTable->Rows->Count; nRowIndex++)
					 {
						 nTotal = nTotal + System::Convert::ToInt32(dtTable->get_Rows()->get_Item(nRowIndex)->get_Item(nColumnIndex)->ToString());						
					 }
					 footerRow[nColumnIndex] =  nTotal.ToString();
				 }
				 dtTable->Rows->Add(footerRow);				
			 }	 
private: System::Void hlkEasyXLS_LinkClicked(System::Object *  sender, System::Windows::Forms::LinkLabelLinkClickedEventArgs *  e)
		 {
			 // Opening the hyperlink target
			 System::Diagnostics::Process::Start(hlkEasyXLS->Text);
		 }

private: System::Void btnExporttoExcel_Click(System::Object *  sender, System::EventArgs *  e)
		 {
			 // Creating an instance of the object that generates excel files
			ExcelDocument *xls = new ExcelDocument();
			
			// Adding a sheet to the Excel Document object
			ExcelWorksheet * xlsWorksheet = new ExcelWorksheet("TimeSheetReport");
			xls->easy_addWorksheet(xlsWorksheet);

			// Adding the image
			xlsWorksheet->easy_addImage(imgEasyXLSlogo->Tag->ToString(), "A1"); 

			// Adding the hyperlink
			xlsWorksheet->easy_addHyperlink(HyperlinkType::URL, hlkEasyXLS->Text, "A5");


			// Creating an instance of the object used to format the cells
			ExcelAutoFormat * xlsAutoFormat = new ExcelAutoFormat(Styles::AUTOFORMAT_EASYXLS1);
			
			// Adding the content of the grid
			xlsWorksheet->easy_insertDataSet(dsSource, 6, 0, xlsAutoFormat, true);

			// Creating the footer
			int  nFooterRowIndex = 6 + dtTable->Rows->Count;
			ExcelTable * xlsTable = xlsWorksheet->easy_getExcelTable();
			xlsTable->easy_getCell(nFooterRowIndex, 0)->setValue("Totals:");
			xlsTable->easy_getCell(nFooterRowIndex, 4)->setValue(String::Concat("=SUM(E8:E" , nFooterRowIndex.ToString() , ")"));
			xlsTable->easy_getCell(nFooterRowIndex, 4)->setDataType(DataType::AUTOMATIC);
			xlsTable->easy_getCell(nFooterRowIndex, 5)->setValue(String::Concat("=SUM(F8:F" , nFooterRowIndex.ToString() , ")"));
			xlsTable->easy_getCell(nFooterRowIndex, 5)->setDataType(DataType::AUTOMATIC);
			xlsTable->easy_getCell(nFooterRowIndex, 6)->setValue(String::Concat("=SUM(G8:G" , nFooterRowIndex.ToString() , ")"));
			xlsTable->easy_getCell(nFooterRowIndex, 6)->setDataType(DataType::AUTOMATIC);
			xlsTable->easy_getCell(nFooterRowIndex, 7)->setValue(String::Concat("=SUM(H8:H" , nFooterRowIndex.ToString() , ")"));
			xlsTable->easy_getCell(nFooterRowIndex, 7)->setDataType(DataType::AUTOMATIC);
			

			// Creating and adding a chart based on the grid's data	
			ExcelChart * xlsChart = new ExcelChart("A20", 600, 300);
			if (chkEstimated->Checked)
				xlsChart->easy_addSeries("=TimeSheetReport!$E$7", "=TimeSheetReport!$E$8:$E$16");
			if (chkRegular->Checked)
				xlsChart->easy_addSeries("=TimeSheetReport!$F$7", "=TimeSheetReport!$F$8:$F$16");
			if (chkOTHours->Checked)
				xlsChart->easy_addSeries("=TimeSheetReport!$G$7", "=TimeSheetReport!$G$8:$G$16");
			if (chkNBHours->Checked)
				xlsChart->easy_addSeries("=TimeSheetReport!$H$7", "=TimeSheetReport!$H$8:$H$16");

			if (chkEstimated->Checked || chkRegular->Checked || chkOTHours->Checked || chkNBHours->Checked)
				xlsChart->easy_setCategoryXAxisLabels("=TimeSheetReport!$D$8:$D$16");
			else
				xlsChart->easy_addSeries("=TimeSheetReport!$D$7", "=TimeSheetReport!$D$8:$D$16");

			xlsWorksheet->easy_addChart(xlsChart);

			// Generating the file
			label1->Text = "Writing file C:\\Samples\\CPlusPlusDotNetApplication.xls.";
			xls->easy_WriteXLSFile("c:\\Samples\\CPlusPlusDotNetApplication.xls");

			//Confirm generation
			String *sError = xls->easy_getError();
			if (sError->Equals(""))
				  label1->Text = String::Concat(label1->Text, "\nFile successfully created.");
			else
				label1->Text = String::Concat("\nError encountered: ", sError);				
				
			// Dispose memory
			xls->Dispose();
		 }

private: System::Void label1_Click(System::Object *  sender, System::EventArgs *  e)
		 {
		 }

};
}


