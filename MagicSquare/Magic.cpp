#include "Magic.h"

Magic::Magic(unsigned deg, unsigned bas, unsigned diff) {//Выделение памяти под квадрат
	degree = deg;
	basis = bas;
	difference = diff;
	square = new unsigned*[degree];
	for (int i = 0; i < degree; ++i)
		square[i] = new unsigned[degree];
}

Magic::~Magic() {//Освобождение памяти
	for (int i = 0; i < degree; ++i)
		delete[] square[i];
	delete[] square;
}

int Magic::checkSum() {//Побочная функция, которая использовалась для тестирования правильности работы. Более не нужна
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

void Magic::buildOddSquare() {//Нечётный случай
	for (int i = 0; i < degree; ++i) {//Изначально всё заполняется нулями
		for (int j = 0; j < degree; ++j) {
			square[i][j] = 0;
		}
	}

	unsigned row = 0, col = degree / 2;//Вспомогательные величины, которые будут "бегать" по методу 

	square[row][col] = basis; //Начальный элемент по методу задаётся первым элементом прогрессии

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
		square[row][col] = basis + difference * (i - 1);//Присваиваем каждой ячейке порядковый номер
	}
}

void Magic::buildEvenSquare() { //Случай с чётным порядком
	unsigned counter = 0; //Переменная для заполнения массива

	for (int i = 0; i < degree; ++i) {
		for (int j = 0; j < degree; ++j) {//Заполнение массива в соответствии с прогрессией
			square[i][j] = basis + counter * difference;
			counter++;
		}
	}

	for (int i = 0; i < degree / 2; ++i) { //top-left
		for (int j = 0; j < degree / 2; ++j) {
			if (i == j)
				square[i][j] = 0;
			if (i == degree / 2 - j - 1) //Побочная диагональ
				square[i][j] = 0;
		}
	}

	for (int i = 0; i < degree / 2; ++i) {//top-right
		for (int j = degree / 2; j < degree; ++j) {
			if (i == j - degree / 2)
				square[i][j] = 0;
			if (i == degree - j - 1) //Побочная диагональ
				square[i][j] = 0;
		}
	}

	for (int i = degree / 2; i < degree; ++i) {//bottom-left
		for (int j = 0; j < degree / 2; ++j) {
			if (i - degree / 2 == j)
				square[i][j] = 0;
			if (i == degree - j - 1) //Побочная диагональ
				square[i][j] = 0;
		}
	}

	for (int i = degree / 2; i < degree; ++i) {//bottom-right
		for (int j = degree / 2; j < degree; ++j) {
			if (i == j)
				square[i][j] = 0;
			if (i == degree + degree / 2 - j - 1) //Побочная диагональ
				square[i][j] = 0;
		}
	}

	counter = basis + difference * (degree - 1);//Заполнение оставшихся элементов квадрата в соответствии с алгоритмом
	for (int i = 0; i < degree; ++i) {
		for (int j = 0; j < degree; ++j) {
			if (square[i][j] == 0)
				square[i][j] = counter;
			counter -= difference;
		}
	}
}


unsigned** Magic::getSquare() {//Передача заполненного квадрата в вызывающий кусок кода
	unsigned** a = new unsigned*[degree];
	for (int i = 0; i < degree; ++i)
		a[i] = new unsigned[degree];

	for (int i = 0; i < degree; ++i)
		for (int j = 0; j < degree; ++j)
			a[i][j] = square[i][j];
	return a;
}