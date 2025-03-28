# Конвертер таблиц из DOCX в Excel

Программа для конвертации таблиц из DOCX файлов в Excel с автоматической обработкой данных.

## Возможности

- Конвертация таблиц из DOCX в Excel
- Автоматическое удаление столбцов A и C
- Нормализация дат в формате ДД.ММ.ГГГГ
- Обработка дат рождения
- Обработка дат окончания
- Перемещение информации о судах
- Графический интерфейс пользователя

## Требования

- Python 3.x
- python-docx
- openpyxl
- tkinter (обычно входит в стандартную библиотеку Python)

## Установка

1. Клонируйте репозиторий:
```bash
git clone https://github.com/ваш-username/docx-to-excel-converter.git
```

2. Установите зависимости:
```bash
pip install python-docx openpyxl
```

## Использование

1. Запустите программу:
```bash
python simple_docx_to_excel.py
```

2. Выберите DOCX файл с помощью кнопки "Обзор..."
3. Нажмите "Обработать" для конвертации
4. Результат будет сохранен в Excel файл в той же директории

## Правила обработки

- Все таблицы из Word переносятся в Excel
- Первая строка удаляется, если во второй ячейке НЕТ даты
- Первая строка сохраняется, если во второй ячейке ЕСТЬ дата
- Столбцы A и C удаляются
- Даты во втором столбце приводятся к формату ДД.ММ.ГГГГ
- Даты рождения в пятом столбце приводятся к формату ДД.ММ.ГГГГ
- Даты окончания в восьмом столбце приводятся к формату ДД.ММ.ГГГГ
- Информация о судах из столбцов D и E перемещается в столбец I

## Лицензия

MIT 