#pragma once
#include "Magic.h"


namespace MagicSquare {

//	using namespace Microsoft::Office::Interop::Excel; // ��� ������������� PIA Office Excel �� Microsoft
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	


	/// <summary>
	/// ������ ��� MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: �������� ��� ������������
			//
		}

	protected:
		/// <summary>
		/// ���������� ��� ������������ �������.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::DataGridView^  dataOutput;
	protected:

	protected:
	private: System::Windows::Forms::Button^  button1;
	private: System::Windows::Forms::TextBox^  txtCapacity;
	private: System::Windows::Forms::Label^  label1;
	private: System::Windows::Forms::MenuStrip^  menuStrip1;
	private: System::Windows::Forms::ToolStripMenuItem^  ����ToolStripMenuItem;

	private: System::Windows::Forms::ToolStripMenuItem^  ����������ToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^  ����������ToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^  ��������������������ToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^  �������ToolStripMenuItem;
	private: System::Windows::Forms::Label^  label2;
	private: System::Windows::Forms::Label^  label3;
	private: System::Windows::Forms::TextBox^  txtBasis;
	private: System::Windows::Forms::TextBox^  txtDiffer;

	private:
		/// <summary>
		/// ������������ ���������� ������������.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// ��������� ����� ��� ��������� ������������ � �� ��������� 
		/// ���������� ����� ������ � ������� ��������� ����.
		/// </summary>
		void InitializeComponent(void)
		{
			this->dataOutput = (gcnew System::Windows::Forms::DataGridView());
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->txtCapacity = (gcnew System::Windows::Forms::TextBox());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
			this->����ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->�������ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->����������ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->����������ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->��������������������ToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->label3 = (gcnew System::Windows::Forms::Label());
			this->txtBasis = (gcnew System::Windows::Forms::TextBox());
			this->txtDiffer = (gcnew System::Windows::Forms::TextBox());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataOutput))->BeginInit();
			this->menuStrip1->SuspendLayout();
			this->SuspendLayout();
			// 
			// dataOutput
			// 
			this->dataOutput->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataOutput->Location = System::Drawing::Point(12, 118);
			this->dataOutput->Name = L"dataOutput";
			this->dataOutput->RowTemplate->Height = 24;
			this->dataOutput->Size = System::Drawing::Size(776, 568);
			this->dataOutput->TabIndex = 0;
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(657, 73);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(131, 39);
			this->button1->TabIndex = 1;
			this->button1->Text = L"���������";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// txtCapacity
			// 
			this->txtCapacity->Location = System::Drawing::Point(362, 25);
			this->txtCapacity->Name = L"txtCapacity";
			this->txtCapacity->Size = System::Drawing::Size(100, 22);
			this->txtCapacity->TabIndex = 2;
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(12, 28);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(217, 17);
			this->label1->TabIndex = 3;
			this->label1->Text = L"������� ����������� ��������";
			// 
			// menuStrip1
			// 
			this->menuStrip1->ImageScalingSize = System::Drawing::Size(20, 20);
			this->menuStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(2) {
				this->����ToolStripMenuItem,
					this->����������ToolStripMenuItem
			});
			this->menuStrip1->Location = System::Drawing::Point(0, 0);
			this->menuStrip1->Name = L"menuStrip1";
			this->menuStrip1->Size = System::Drawing::Size(800, 28);
			this->menuStrip1->TabIndex = 4;
			this->menuStrip1->Text = L"menuStrip1";
			// 
			// ����ToolStripMenuItem
			// 
			this->����ToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) { this->�������ToolStripMenuItem });
			this->����ToolStripMenuItem->Name = L"����ToolStripMenuItem";
			this->����ToolStripMenuItem->Size = System::Drawing::Size(57, 24);
			this->����ToolStripMenuItem->Text = L"����";
			// 
			// �������ToolStripMenuItem
			// 
			this->�������ToolStripMenuItem->Name = L"�������ToolStripMenuItem";
			this->�������ToolStripMenuItem->Size = System::Drawing::Size(141, 26);
			this->�������ToolStripMenuItem->Text = L"�������";
			this->�������ToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::�������ToolStripMenuItem_Click);
			// 
			// ����������ToolStripMenuItem
			// 
			this->����������ToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(2) {
				this->����������ToolStripMenuItem,
					this->��������������������ToolStripMenuItem
			});
			this->����������ToolStripMenuItem->Name = L"����������ToolStripMenuItem";
			this->����������ToolStripMenuItem->Size = System::Drawing::Size(116, 24);
			this->����������ToolStripMenuItem->Text = L"� ���������";
			// 
			// ����������ToolStripMenuItem
			// 
			this->����������ToolStripMenuItem->Name = L"����������ToolStripMenuItem";
			this->����������ToolStripMenuItem->Size = System::Drawing::Size(263, 26);
			this->����������ToolStripMenuItem->Text = L"����������";
			this->����������ToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::����������ToolStripMenuItem_Click);
			// 
			// ��������������������ToolStripMenuItem
			// 
			this->��������������������ToolStripMenuItem->Name = L"��������������������ToolStripMenuItem";
			this->��������������������ToolStripMenuItem->Size = System::Drawing::Size(263, 26);
			this->��������������������ToolStripMenuItem->Text = L"���������� � ���������";
			this->��������������������ToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::��������������������ToolStripMenuItem_Click);
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Location = System::Drawing::Point(12, 58);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(301, 17);
			this->label2->TabIndex = 5;
			this->label2->Text = L"������� ����� �������������� ����������";
			// 
			// label3
			// 
			this->label3->AutoSize = true;
			this->label3->Location = System::Drawing::Point(12, 84);
			this->label3->Name = L"label3";
			this->label3->Size = System::Drawing::Size(323, 17);
			this->label3->TabIndex = 6;
			this->label3->Text = L"������� �������� �������������� ����������";
			// 
			// txtBasis
			// 
			this->txtBasis->Location = System::Drawing::Point(362, 53);
			this->txtBasis->Name = L"txtBasis";
			this->txtBasis->Size = System::Drawing::Size(100, 22);
			this->txtBasis->TabIndex = 7;
			this->txtBasis->Text = L"1";
			// 
			// txtDiffer
			// 
			this->txtDiffer->Location = System::Drawing::Point(362, 79);
			this->txtDiffer->Name = L"txtDiffer";
			this->txtDiffer->Size = System::Drawing::Size(100, 22);
			this->txtDiffer->TabIndex = 8;
			this->txtDiffer->Text = L"1";
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(8, 16);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(800, 698);
			this->Controls->Add(this->txtDiffer);
			this->Controls->Add(this->txtBasis);
			this->Controls->Add(this->label3);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->txtCapacity);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->dataOutput);
			this->Controls->Add(this->menuStrip1);
			this->MainMenuStrip = this->menuStrip1;
			this->Name = L"MyForm";
			this->Text = L"���������� ���������� ���������";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataOutput))->EndInit();
			this->menuStrip1->ResumeLayout(false);
			this->menuStrip1->PerformLayout();
			this->ResumeLayout(false);
			this->PerformLayout();

		}

		bool checkIfAllValuesAreNumbers(String^ checked) {//��������, ��� �� ������� ������ - �����
			for (int i = 0; i < checked->Length; ++i) {
				if ((int)checked[i] < 48 || (int)checked[i] > 57) {
					return false;
				}
			}
			return true;
		}


#pragma endregion
	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {
		if (txtCapacity->Text == "") {//�������� �� �����������
			MessageBox::Show("�� �� ����� �����������!");
			return;
		}
		
		if(!checkIfAllValuesAreNumbers(txtCapacity->Text) || !checkIfAllValuesAreNumbers(txtDiffer->Text) || !checkIfAllValuesAreNumbers(txtBasis->Text)) {
			MessageBox::Show("������� ������ �����!");//���� ����� �� ������ �����, �� � ��������� �������
			return;
		}

		if (System::Convert::ToInt16(txtDiffer->Text) == 0) {//���� ������� �������������� ���������� ����� ����
			MessageBox::Show("�������, � ������� ��� ����� ���������, �������, ����������, �� ��� �� ���������");
			return;
		}
		//����������-��������� ����������� ��������, ������ � ������ �������������� ����������
		int size = System::Convert::ToInt16(txtCapacity->Text);
		int basis = System::Convert::ToDouble(txtBasis->Text);
		int differ = System::Convert::ToDouble(txtDiffer->Text);
		
		if (size < 0) {//���� ������������ ��� ������������� �����������
			MessageBox::Show("������� ������������� �����������!");
			return;
		}

		if ((size < 5) && (size % 2 == 0)) {//���� ������� ������� �� ����� ������
			MessageBox::Show("������� ������ �����������, ������� 6 �� ����� ������");
			return;
		}

		Magic* m = new Magic(size, basis, differ);//�������� ���������� ������
		m->build();

		unsigned** square = m->getSquare();//��������� ������������ ��������
		//���������� �������� dataGridView
		dataOutput->RowCount = size;
		dataOutput->ColumnCount = size;

		for (int i = 0; i < size; ++i)//���������� dataGridView
			for (int j = 0; j < size; ++j)
				dataOutput->Rows[i]->Cells[j]->Value = System::Convert::ToString(square[i][j]);
		
		m->~Magic();//������������ ������
		for (int i = 0; i < size; ++i)
			delete[] square[i];
		delete[] square;
	}
			 //������ � ����
	private: System::Void ��������������������ToolStripMenuItem_Click(System::Object^  sender, System::EventArgs^  e) {
		MessageBox::Show("��������� - ������� ������ ��-65-17 ����� ������ �������������. ���������, � ���������, ���� ������� ������������ :(");
	}
private: System::Void ����������ToolStripMenuItem_Click(System::Object^  sender, System::EventArgs^  e) {
	MessageBox::Show("��������� ������������� ��� ���������� ���������� ���������. ���������� ���������� �������, � �������� ����� ����� �� ���� �������, �������� � ���������� ����� ����� �����. ��������� �������� � ����������, ����������� ������� ����� 3 ��� ������ 4-�");
}
private: System::Void �������ToolStripMenuItem_Click(System::Object^  sender, System::EventArgs^  e) {//�������� ��������� ����� ����
	Application::Exit();
}
};
} 
