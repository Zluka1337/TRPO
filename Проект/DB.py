import sqlite3

class DB:
    def __init__(self):
        self.conn = sqlite3.connect('resident.db')
        self.c = self.conn.cursor()

        self.c.execute('''
            CREATE TABLE IF NOT EXISTS anketa (
                id_anketa INTEGER PRIMARY KEY AUTOINCREMENT,
                id_citizen TEXT NOT NULL,
                date_of_birth TEXT NOT NULL,
                citizenship TEXT NOT NULL,
                id_gender INTEGER NOT NULL,
                home_address TEXT NOT NULL,
                place_of_birth TEXT NOT NULL,
                inn TEXT,
                insurance_number TEXT NOT NULL,
                phone_number TEXT NOT NULL,
                marital_status TEXT NOT NULL,
                additional_info TEXT,
                employer TEXT,
                polling_station_number TEXT
            )
        ''')
        self.conn.commit()

        self.c.execute('''
            CREATE TABLE IF NOT EXISTS citizen (
                id_citizen INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL
            )
        ''')
        self.conn.commit()

        self.c.execute('''
            CREATE TABLE IF NOT EXISTS gender (
                id_gender INTEGER PRIMARY KEY AUTOINCREMENT,
                name_gender TEXT NOT NULL
            )
        ''')
        self.conn.commit()

# Пример использования
if __name__ == '__main__':
    db = DB()
