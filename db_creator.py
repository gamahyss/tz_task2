import sqlite3

def create_database(db_path='files.db'):
    '''функция для создания базы данных которая хранит содержимое файлов из хранилища'''
    
    with sqlite3.connect(db_path) as conn: # подключаемся к базе данных
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS files (
                file_path TEXT,
                file_name TEXT,
                extension TEXT,
                content TEXT
            )
        ''') # создаём таблицу если её ещё не существует
        conn.commit() # фиксируем изменения

if __name__ == "__main__":
    create_database()