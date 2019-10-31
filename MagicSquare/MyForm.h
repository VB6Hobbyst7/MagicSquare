#pragma once
#include "Magic.h"


namespace MagicSquare {

//	using namespace Microsoft::Office::Interop::Excel; // для использования PIA Office Excel от Microsoft
	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	


	/// <summary>
	/// Сводка для MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
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
	private: System::Windows::Forms::ToolStripMenuItem^  файлToolStripMenuItem;

	private: System::Windows::Forms::ToolStripMenuItem^  оПрограммеToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^  назначениеToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^  информацияОСоздаталеToolStripMenuItem;
	private: System::Windows::Forms::ToolStripMenuItem^  закрытьToolStripMenuItem;
	private: System::Windows::Forms::Label^  label2;
	private: System::Windows::Forms::Label^  label3;
	private: System::Windows::Forms::TextBox^  txtBasis;
	private: System::Windows::Forms::TextBox^  txtDiffer;

	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->dataOutput = (gcnew System::Windows::Forms::DataGridView());
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->txtCapacity = (gcnew System::Windows::Forms::TextBox());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->menuStrip1 = (gcnew System::Windows::Forms::MenuStrip());
			this->файлToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->закрытьToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->оПрограммеToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->назначениеToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
			this->информацияОСоздаталеToolStripMenuItem = (gcnew System::Windows::Forms::ToolStripMenuItem());
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
			this->button1->Text = L"Построить";
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
			this->label1->Text = L"Введите размерность квадрата";
			// 
			// menuStrip1
			// 
			this->menuStrip1->ImageScalingSize = System::Drawing::Size(20, 20);
			this->menuStrip1->Items->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(2) {
				this->файлToolStripMenuItem,
					this->оПрограммеToolStripMenuItem
			});
			this->menuStrip1->Location = System::Drawing::Point(0, 0);
			this->menuStrip1->Name = L"menuStrip1";
			this->menuStrip1->Size = System::Drawing::Size(800, 28);
			this->menuStrip1->TabIndex = 4;
			this->menuStrip1->Text = L"menuStrip1";
			// 
			// файлToolStripMenuItem
			// 
			this->файлToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(1) { this->закрытьToolStripMenuItem });
			this->файлToolStripMenuItem->Name = L"файлToolStripMenuItem";
			this->файлToolStripMenuItem->Size = System::Drawing::Size(57, 24);
			this->файлToolStripMenuItem->Text = L"Файл";
			// 
			// закрытьToolStripMenuItem
			// 
			this->закрытьToolStripMenuItem->Name = L"закрытьToolStripMenuItem";
			this->закрытьToolStripMenuItem->Size = System::Drawing::Size(141, 26);
			this->закрытьToolStripMenuItem->Text = L"Закрыть";
			this->закрытьToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::закрытьToolStripMenuItem_Click);
			// 
			// оПрограммеToolStripMenuItem
			// 
			this->оПрограммеToolStripMenuItem->DropDownItems->AddRange(gcnew cli::array< System::Windows::Forms::ToolStripItem^  >(2) {
				this->назначениеToolStripMenuItem,
					this->информацияОСоздаталеToolStripMenuItem
			});
			this->оПрограммеToolStripMenuItem->Name = L"оПрограммеToolStripMenuItem";
			this->оПрограммеToolStripMenuItem->Size = System::Drawing::Size(116, 24);
			this->оПрограммеToolStripMenuItem->Text = L"О программе";
			// 
			// назначениеToolStripMenuItem
			// 
			this->назначениеToolStripMenuItem->Name = L"назначениеToolStripMenuItem";
			this->назначениеToolStripMenuItem->Size = System::Drawing::Size(263, 26);
			this->назначениеToolStripMenuItem->Text = L"Назначение";
			this->назначениеToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::назначениеToolStripMenuItem_Click);
			// 
			// информацияОСоздаталеToolStripMenuItem
			// 
			this->информацияОСоздаталеToolStripMenuItem->Name = L"информацияОСоздаталеToolStripMenuItem";
			this->информацияОСоздаталеToolStripMenuItem->Size = System::Drawing::Size(263, 26);
			this->информацияОСоздаталеToolStripMenuItem->Text = L"Информация о создатале";
			this->информацияОСоздаталеToolStripMenuItem->Click += gcnew System::EventHandler(this, &MyForm::информацияОСоздаталеToolStripMenuItem_Click);
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Location = System::Drawing::Point(12, 58);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(301, 17);
			this->label2->TabIndex = 5;
			this->label2->Text = L"Введите базис арифметической прогрессии";
			// 
			// label3
			// 
			this->label3->AutoSize = true;
			this->label3->Location = System::Drawing::Point(12, 84);
			this->label3->Name = L"label3";
			this->label3->Size = System::Drawing::Size(323, 17);
			this->label3->TabIndex = 6;
			this->label3->Text = L"Введите разность арифметической прогрессии";
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
			this->Text = L"Построение магических квадратов";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataOutput))->EndInit();
			this->menuStrip1->ResumeLayout(false);
			this->menuStrip1->PerformLayout();
			this->ResumeLayout(false);
			this->PerformLayout();

		}

		bool checkIfAllValuesAreNumbers(String^ checked) {//Проверка, все ли символы строки - числа
			for (int i = 0; i < checked->Length; ++i) {
				if ((int)checked[i] < 48 || (int)checked[i] > 57) {
					return false;
				}
			}
			return true;
		}


#pragma endregion
	private: System::Void button1_Click(System::Object^  sender, System::EventArgs^  e) {
		if (txtCapacity->Text == "") {//Проверка на размерность
			MessageBox::Show("Вы не ввели размерность!");
			return;
		}
		
		if(!checkIfAllValuesAreNumbers(txtCapacity->Text) || !checkIfAllValuesAreNumbers(txtDiffer->Text) || !checkIfAllValuesAreNumbers(txtBasis->Text)) {
			MessageBox::Show("Вводите только цифры!");//Если ввели не только цифры, но и сторонние символы
			return;
		}

		if (System::Convert::ToInt16(txtDiffer->Text) == 0) {//Если разница арифметической прогрессии равна нулю
			MessageBox::Show("Квадрат, в котором все числа одинаковы, конечно, магический, но это не интересно");
			return;
		}
		//Переменные-хранители размерности квадрата, базиса и дельты арифметической прогрессии
		int size = System::Convert::ToInt16(txtCapacity->Text);
		int basis = System::Convert::ToDouble(txtBasis->Text);
		int differ = System::Convert::ToDouble(txtDiffer->Text);
		
		if (size < 0) {//Если пользователь ввёл отрицательную размерность
			MessageBox::Show("Введите положительную размерность!");
			return;
		}

		if ((size < 5) && (size % 2 == 0)) {//Если строить квадрат не имеет смысла
			MessageBox::Show("Вводить чётную размерность, меньшую 6 не имеет смысла");
			return;
		}

		Magic* m = new Magic(size, basis, differ);//Создание экземпляра класса
		m->build();

		unsigned** square = m->getSquare();//Получение построенного квадрата
		//Выделениие размеров dataGridView
		dataOutput->RowCount = size;
		dataOutput->ColumnCount = size;

		for (int i = 0; i < size; ++i)//Заполнение dataGridView
			for (int j = 0; j < size; ++j)
				dataOutput->Rows[i]->Cells[j]->Value = System::Convert::ToString(square[i][j]);
		
		m->~Magic();//Освобождение памяти
		for (int i = 0; i < size; ++i)
			delete[] square[i];
		delete[] square;
	}
			 //Работа с меню
	private: System::Void информацияОСоздаталеToolStripMenuItem_Click(System::Object^  sender, System::EventArgs^  e) {
		MessageBox::Show("Создатель - студент группы ИЭ-65-17 Гасин Михаил Александрович. Программа, к сожалению, была создана безвозмездно :(");
	}
private: System::Void назначениеToolStripMenuItem_Click(System::Object^  sender, System::EventArgs^  e) {
	MessageBox::Show("Программа предназначена для построения магических квадратов. Магическим называется квадрат, у которого суммы чисел по всем строкам, столбцам и диагоналям равны между собой. Программа работает с квадратами, размерность которых равна 3 или больше 4-х");
}
private: System::Void закрытьToolStripMenuItem_Click(System::Object^  sender, System::EventArgs^  e) {//Закрытие программы через меню
	Application::Exit();
}
};
} 
