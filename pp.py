import sqlite3

class ResidentDB:
    def __init__(self, db_name='resident_db.db'):
        self.conn = sqlite3.connect(db_name)
        self.c = self.conn.cursor()
        self.create_tables()

    def create_tables(self):
        # Создание таблицы для анкетных данных жителей
        self.c.execute('''
            CREATE TABLE IF NOT EXISTS residents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                date_of_birth TEXT NOT NULL,
                citizenship TEXT NOT NULL,
                gender TEXT NOT NULL,
                home_address TEXT NOT NULL,
                place_of_birth TEXT NOT NULL,
                inn TEXT ,
                insurance_number TEXT NOT NULL,
                phone_number TEXT NOT NULL,
                marital_status TEXT NOT NULL,
                additional_info TEXT,
                employer TEXT,
                polling_station_number TEXT
            )
        ''')
        self.conn.commit()

    def insert_resident(self, data):
        # Вставка данных жителя
        query = '''
            INSERT INTO residents (
                full_name, date_of_birth, citizenship, gender, home_address, place_of_birth,
                inn, insurance_number, phone_number, marital_status, additional_info,
                employer, polling_station_number
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        '''
        self.c.execute(query, data)
        self.conn.commit()

    def update_resident(self, resident_id, data):
        # Обновление данных жителя
        query = '''
            UPDATE residents
            SET full_name=?, date_of_birth=?, citizenship=?, gender=?, home_address=?, place_of_birth=?,
                inn=?, insurance_number=?, phone_number=?, marital_status=?, additional_info=?,
                employer=?, polling_station_number=?
            WHERE id=?
        '''
        self.c.execute(query, (*data, resident_id))
        self.conn.commit()

    def delete_resident(self, resident_id):
        # Удаление данных жителя по ID
        query = 'DELETE FROM residents WHERE id=?'
        self.c.execute(query, (resident_id,))
        self.conn.commit()

    def get_residents(self):
        # Получение списка всех жителей
        query = 'SELECT * FROM residents'
        self.c.execute(query)
        return self.c.fetchall()

    def get_resident_by_id(self, resident_id):
        # Получение данных жителя по ID
        query = 'SELECT * FROM residents WHERE id=?'
        self.c.execute(query, (resident_id,))
        return self.c.fetchone()

# Пример использования
if __name__ == '__main__':
    db = ResidentDB()

    # Вставка нового жителя
    new_resident_data = (
        "Иванов Иван Иванович",
        "01.01.1990",
        "Россия",
        "Мужской",
        "г. Москва, ул. Примерная, д. 123",
        "г. Москва",
        "123456789012",
        "987654321098",
        "+7 (123) 456-78-90",
        "Женат/Замужем",
        "Дополнительные сведения",
        "ООО Работодатель",
        "123"
    )
    db.insert_resident(new_resident_data)

    # Обновление данных жителя по ID
    resident_id_to_update = 1
    updated_data = (
        "Иванов Иван Петрович",
        "02.02.1995",
        "Россия",
        "Мужской",
        "г. Москва, ул. Примерная, д. 124",
        "г. Москва",
        "987654321012",
        "123456789098",
        "+7 (123) 456-78-91",
        "Холост/Не замужем",
        "Новые сведения",
        "ООО Новый Работодатель",
        "124"
    )
    db.update_resident(resident_id_to_update, updated_data)

    # Получение списка всех жителей и вывод на экран
    residents = db.get_residents()
    for resident in residents:
        print(resident)

    # Удаление жителя по ID
    resident_id_to_delete = 2
    db.delete_resident(resident_id_to_delete)
