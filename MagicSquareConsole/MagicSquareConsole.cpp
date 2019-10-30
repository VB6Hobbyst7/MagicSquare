#include "pch.h"
#include <iostream>

class Magic {
private:
	unsigned degree;        //Порядок квадрата
	unsigned** square;      //Магический квадрат
public:
	Magic(unsigned deg);//.............................Конструктор
	~Magic();//........................................Деструктор
	void buildOddSquare(); //..........................Заполнение для магического квадрата нечётного порядка
	void buildEvenSquare();//..........................Заполнение для магического квадрата чётного порядка
	int checkSum();        //..........................Если квадрат является магическим, возвращает его сумму. Если нет, возвращает 0
	void build();          //..........................Для чётного порядка - вызов buildEvenSquare, для нечётного - buildOddSquare
	void show();           //..........................Вывод квадрата на экран
};

Magic::Magic(unsigned deg) {//Выделение памяти под квадрат
	degree = deg;
	square = new unsigned*[degree];
	for (int i = 0; i < degree; ++i)
		square[i] = new unsigned[degree];
}

Magic::~Magic() {//Освобождение памяти
	for (int i = 0; i < degree; ++i)
		delete[] square[i];
	delete[] square;
}

int Magic::checkSum() {
	int sum = 0, sum1 = 0;

	for (int i = 0; i < degree; ++i)
		sum += square[i][i];

	for (int i = 0; i < degree; ++i)
		sum1 += square[i][degree - i + 1];

	if (sum1 != sum)
		return 0;

	for (int i = 0; i < degree; ++i) {
		sum1 = 0;
		int sum2 = 0;

		for (int j = 0; j < degree; ++j) {
			sum1 += square[i][j];
			sum2 += square[j][i];
		}

		if (sum1 != sum)
			return 0;

		if (sum2 != sum)
			return 0;
	}
	return sum;
}

void Magic::build() {//Вызов одной из двух функций
	if (degree % 2 == 0)
		buildEvenSquare();
	else
		buildOddSquare();
}

void Magic::show() {//Вывод на экран
	for (int i = 0; i < degree; ++i) {
		for (int j = 0; j < degree; ++j) {
			std::cout << square[i][j] << "\t";
		}
		std::cout << std::endl;
	}
}

void Magic::buildOddSquare() {//Нечётный случай
	for (int i = 0; i < degree; ++i) {//Изначально всё заполняется нулями
		for (int j = 0; j < degree; ++j) {
			square[i][j] = 0;
		}
	}

	unsigned row = 0, col = degree / 2;//Вспомогательные величины, которые будут "бегать" по методу 

	square[row][col] = 1; //Начальный элемент по методу задаётся единичным

	for (int i = 2; i <= degree * degree; ++i) {//Дальше бегаем по квадрату по диагонали. Если ушли за пределы, возвращаемся снизу
		int rowT = row - 1;
		int colT = col + 1;
		if (rowT < 0)
			rowT += degree;
		if (colT >= degree)
			colT -= degree;
		if (square[rowT][colT]) {
			if (++row >= degree)
				row -= 3;
		}
		else {
			row = rowT;
			col = colT;
		}
		square[row][col] = i;//Присваиваем каждой ячейке порядковый номер
	}
}

void Magic::buildEvenSquare() { //Случай с чётным порядком
	for (int i = 0; i < degree; ++i)
		for (int j = 0; j < degree; ++j)
			square[i][j] = degree * i + j + 1;

	for (int i = 0; i < degree / 2; ++i) { //top-left
		for (int j = 0; j < degree / 2; ++j) {
			if (i == j)
				square[i][j] = 0;
			if (i == degree / 2 - j - 1) //РЈСЃР»РѕРІРёРµ РїРѕР±РѕС‡РЅРѕР№ РґРёР°РіРѕРЅР°Р»Рё
				square[i][j] = 0;
		}
	}

	for (int i = 0; i < degree / 2; ++i) {//top-right
		for (int j = degree / 2; j < degree; ++j) {
			if (i == j - degree / 2)
				square[i][j] = 0;
			if (i == degree - j - 1) //РЈСЃР»РѕРІРёРµ РїРѕР±РѕС‡РЅРѕР№ РґРёР°РіРѕРЅР°Р»Рё
				square[i][j] = 0;
		}
	}

	for (int i = degree / 2; i < degree; ++i) {//bottom-left
		for (int j = 0; j < degree / 2; ++j) {
			if (i - degree / 2 == j)
				square[i][j] = 0;
			if (i == degree - j - 1) //РЈСЃР»РѕРІРёРµ РїРѕР±РѕС‡РЅРѕР№ РґРёР°РіРѕРЅР°Р»Рё
				square[i][j] = 0;
		}
	}

	for (int i = degree / 2; i < degree; ++i) {//bottom-right
		for (int j = degree / 2; j < degree; ++j) {
			if (i == j)
				square[i][j] = 0;
			if (i == degree + degree / 2 - j - 1) //РЈСЃР»РѕРІРёРµ РїРѕР±РѕС‡РЅРѕР№ РґРёР°РіРѕРЅР°Р»Рё
				square[i][j] = 0;
		}
	}

	std::cout << std::endl;
	this->show();
	std::cout << std::endl;

	unsigned counter = degree * degree;
	for (int i = 0; i < degree; ++i) {
		for (int j = 0; j < degree; ++j) {
			if (square[i][j] == 0)
				square[i][j] = counter;
			counter--;
		}
	}
}

int main()
{
	int n = 0;
	std::cout << "Enter n" << std::endl;
	while (n <= 1) {
		try {
			std::cin >> n;
			if (n <= 0)
				throw(123);
			if (n == 1)
				throw('a');
		}
		catch (int a) {
			std::cout << "Enter a positive number" << std::endl;
		}
		catch (char c) {
			std::cout << "Enter a number that is different from 1" << std::endl;
		}
	}
	Magic* m = new Magic(n, 1, 0);
	m->build();
	m->show();
	std::cout << m->checkSum();
}