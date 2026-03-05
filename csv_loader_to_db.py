import sqlite3
import csv

def load_csv_to_db(csv_file, db_file='files.db'):
    '''загружает данные из csv-файла в таблицу files'''

    conn = sqlite3.connect(db_file) # подключаемся к базе данных
    cursor = conn.cursor()
    with open(csv_file, 'r', encoding='utf-8') as f: # открываем csv-файл в режиме чтения с кодировкой utf-8
        reader = csv.reader(f, delimiter=',', quotechar='"') # создаём объект который будет считывать содержимое
        header = next(reader, None) # пропускаем заголовок
        if header is None: # если заголовок пустой, значит и файл пустой, тогда останавливаем функцию
            print("CSV-файл пуст.")
            return
        expected_header = ['file_path', 'file_name', 'extension', 'content']
        if header != expected_header: # проверяем соответствие заголовка, если он не тот, значит csv не тот который нам нужен
            return
        for row in reader: # перебираем строки по одной
            if len(row) != 4: # если в строке не 4 элемента, то мы не можем вставить её в бд
                continue
            file_path, file_name, extension, content = row # находим из строки значения всех колонок
            cursor.execute('''
                INSERT INTO files (file_path, file_name, extension, content)
                VALUES (?, ?, ?, ?)
            ''', (file_path, file_name, extension, content)) # вставляем полученные значения в бд
    conn.commit() # сохраняем изменения и закрываем соединение
    print(f"Данные из '{csv_file}' успешно загружены в '{db_file}'.")
    conn.close()

if __name__ == "__main__":
    '''главный блок программы'''

    csv_file = input('Введите полный путь к csv-файлу: ') # опрашиваем у пользователя путь к файлу и добляем содержимое этого файла в бд
    load_csv_to_db(csv_file)