import sqlite3, xlsxwriter, sys, os
import tkinter as tk
from tkinter import ttk
from DB import DB
import pandas as pd
from tkinter.messagebox import showerror, showinfo

CITIZEN_HEADERS = ["№", "Полное имя"]
GENDER_HEADERS = ["№", "Название"]
ANKETA_HEADERS = ["№", "Полное имя гражданина", "Дата рождения", "Гражданство", "гендер", 
                  "Адрес", "Место рождения", "ИНН", "Номер страховки", "Номер телефона",
                  "Материальный статус", "Дополнительная информация", "Работодатель", 
                  "номер избирательного участка"]

class WindowMain(tk.Tk):
    def __init__(self):
        super().__init__()
        self.last_headers = None

        # Создание фрейма для отображения таблицы
        self.table_frame = tk.Frame(self, width=700, height=400)
        self.table_frame.grid(row=0, column=0, padx=5, pady=5)

        # Загружаем изображение и отображаем его в виджете Label
        label = tk.Label(self.table_frame, text='Таблица не открыта', font=("Calibri", 40))
        label.place(relwidth=1, relheight=1)

        # Создание меню
        self.menu_bar = tk.Menu(self, background='#555', foreground='white')

        # Меню "Файл"
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="Выход", command=self.quit)
        self.menu_bar.add_cascade(label="Файл", menu=file_menu)

        # Меню "Справочники"
        references_menu = tk.Menu(self.menu_bar, tearoff=0)
        references_menu.add_command(label="Граждане", 
                                    command=lambda: self.show_table("SELECT * FROM citizen", CITIZEN_HEADERS))
        references_menu.add_command(label="Гендеры", 
                                    command=lambda: self.show_table("SELECT * FROM gender", GENDER_HEADERS))
        self.menu_bar.add_cascade(label="Справочники", menu=references_menu)

        # Меню "Таблицы"
        tables_menu = tk.Menu(self.menu_bar, tearoff=0)
        tables_menu.add_command(label="Анкета", command=lambda: self.show_table('''
                    SELECT anketa.id_anketa, citizen.full_name, anketa.date_of_birth, anketa.citizenship, gender.name_gender, 
                           anketa.home_address, anketa.place_of_birth, anketa.inn, anketa.insurance_number, anketa.phone_number, 
                           anketa.marital_status, anketa.additional_info, anketa.employer, anketa.polling_station_number
                    FROM anketa                                                  
                    JOIN citizen ON anketa.id_citizen = citizen.id_citizen
                    JOIN gender ON anketa.id_gender = gender.id_gender
        ''', ANKETA_HEADERS))
        self.menu_bar.add_cascade(label="Таблицы", menu=tables_menu)

        # Меню "Отчёты"
        reports_menu = tk.Menu(self.menu_bar, tearoff=0)
        reports_menu.add_command(label="Создать Отчёт", command=self.to_xlsx)
        self.menu_bar.add_cascade(label="Отчёты", menu=reports_menu)

        # Меню "Сервис"
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        help_menu.add_command(label="Руководство пользователя")
        help_menu.add_command(label="O программе")
        self.menu_bar.add_cascade(label="Сервис", menu=help_menu)

        # Установка меню в главное окно
        self.config(menu=self.menu_bar)

        btn_width = 15
        pad = 5

        # Создание кнопок и виджетов для поиска и редактирования данных
        btn_frame = tk.Frame(self)
        btn_frame.grid(row=0, column=1)
        tk.Button(btn_frame, text="добавить", width=btn_width, command=self.add).pack(pady=pad)
        tk.Button(btn_frame, text="удалить", width=btn_width, command=self.delete).pack(pady=pad)
        tk.Button(btn_frame, text="изменить", width=btn_width, command=self.change).pack(pady=pad)

        search_frame = tk.Frame(self)
        search_frame.grid(row=1, column=0, pady=pad)
        self.search_entry = tk.Entry(search_frame, width=30)
        self.search_entry.grid(row=0, column=0, padx=pad)
        tk.Button(search_frame, text="Поиск", command=self.search).grid(row=0, column=1, padx=pad)
        tk.Button(search_frame, text="Искать далее", command=self.search_next).grid(row=0, column=2, padx=pad)
        tk.Button(search_frame, text="Сброс", command=self.reset_search).grid(row=0, column=3, padx=pad)

    def search_in_table(self, table, search_terms, start_item=None):
        table.selection_remove(table.selection())  # Сброс предыдущего выделения

        items = table.get_children('')
        start_index = items.index(start_item) + 1 if start_item else 0

        for item in items[start_index:]:
            values = table.item(item, 'values')
            for term in search_terms:
                if any(term.lower() in str(value).lower() for value in values):
                    table.selection_add(item)
                    table.focus(item)
                    table.see(item)
                    return item  # Возвращаем найденный элемент

    def reset_search(self):
        if self.last_headers:
            self.table.selection_remove(self.table.selection())
        self.search_entry.delete(0, 'end')

    def search(self):
        if self.last_headers:
            self.current_item = self.search_in_table(self.table, self.search_entry.get().split(','))

    def search_next(self):
        if self.last_headers:
            if self.current_item:
                self.current_item = self.search_in_table(self.table, self.search_entry.get().split(','), start_item=self.current_item)
    
    def to_xlsx(self):
        dir = sys.path[0] + "\\export"
        os.makedirs(dir, exist_ok=True)
        path = dir + "\\resident.xlsx"

        # Подключение к базе данных SQLite
        conn = sqlite3.connect('resident.db')
        cursor = conn.cursor()
        # Получите данные из базы данных
        cursor.execute(self.last_sql_query)
        data = cursor.fetchall()
        # Создайте DataFrame из данных
        df = pd.DataFrame(data, columns=self.last_headers)
        # Создайте объект writer для записи данных в Excel
        writer = pd.ExcelWriter(path, engine='xlsxwriter')
        # Запишите DataFrame в файл Excel
        df.to_excel(writer, 'Лист 1', index=False)
        # Сохраните результат
        writer.close()

        showinfo(title="Успешно", message=f"Данные экспортированы в {path}")
    
    def add(self):
        if self.last_headers == GENDER_HEADERS:
            WindowDirectory("add", ("Гендеры", "gender", "id_gender", "name_gender"))
        elif self.last_headers == CITIZEN_HEADERS:
            WindowDirectory("add", ("Граждане", "citizen", "id_citizen", "full_name"))
        elif self.last_headers == ANKETA_HEADERS:
            WindowAnketa("add")
        else: return

        self.withdraw()

    def delete(self):
        if self.last_headers:
            select_item = self.table.selection()
            if select_item:
                item_data = self.table.item(select_item[0])["values"]
            else:
                showerror(title="Ошибка", message="He выбранна запись")
                return
        else:
            return

        if self.last_headers == GENDER_HEADERS:
            WindowDirectory("delete", ("Гендеры", "gender", "id_gender", "name_gender"), item_data)
        elif self.last_headers == CITIZEN_HEADERS:
            WindowDirectory("delete", ("Граждане", "citizen", "id_citizen", "full_name"), item_data)
        elif self.last_headers == ANKETA_HEADERS:
            WindowAnketa("delete", item_data)
        else: return

        self.withdraw()

    def change(self):
        if self.last_headers:
            select_item = self.table.selection()
            if select_item:
                item_data = self.table.item(select_item[0])["values"]
            else:
                showerror(title="Ошибка", message="He выбранна запись")
                return
        else:
            return

        if self.last_headers == GENDER_HEADERS:
            WindowDirectory("change", ("Гендеры", "gender", "id_gender", "name_gender"), item_data)
        elif self.last_headers == CITIZEN_HEADERS:
            WindowDirectory("change", ("Граждане", "citizen", "id_citizen", "full_name"), item_data)
        elif self.last_headers == ANKETA_HEADERS:
             WindowAnketa("change", item_data)
        else: return
        
        self.withdraw()
    
    def show_table(self, sql_query, headers = None):
        # Очистка фрейма перед отображением новых данных
        for widget in self.table_frame.winfo_children(): widget.destroy()

        # Подключение к базе данных SQLite
        conn = sqlite3.connect("resident.db")
        cursor = conn.cursor()

        # Выполнение SQL-запроса
        cursor.execute(sql_query)
        self.last_sql_query = sql_query

        # Получение заголовков таблицы и данных
        if headers == None: # если заголовки не были переданы используем те что в БД
            table_headers = [description[0] for description in cursor.description]
        else: # иначе используем те что передали
            table_headers = headers
            self.last_headers = headers
        table_data = cursor.fetchall()

        # Закрытие соединения с базой данных
        conn.close()
            
        canvas = tk.Canvas(self.table_frame, width=865, height=480)
        canvas.pack(fill="both", expand=True)

        x_scrollbar = ttk.Scrollbar(self.table_frame, orient="horizontal", command=canvas.xview)
        x_scrollbar.pack(side="bottom", fill="x")

        canvas.configure(xscrollcommand=x_scrollbar.set)

        self.table = ttk.Treeview(self.table_frame, columns=table_headers, show="headings", height=23)
        for header in table_headers: 
            self.table.heading(header, text=header)
            self.table.column(header, width=len(header) * 10 + 15) # установка ширины столбца исходя длины его заголовка
        for row in table_data: self.table.insert("", "end", values=row)

        canvas.create_window((0, 0), window=self.table, anchor="nw")

        self.table.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))
    
    def update_table(self):
        self.show_table(self.last_sql_query, self.last_headers)

class WindowDirectory(tk.Toplevel):
    def __init__(self, operation: str, table_info: tuple[str, str, str, str], data = None):
        super().__init__()

        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())
        if data:
            self.id = data[0] # id в таблице спрвочнике
            self.value = data[1] # значение по id
        
        self.table_name_user = table_info[0]
        self.table_name_db = table_info[1]
        self.field_id = table_info[2]
        self.field_name = table_info[3]

        if operation == "add":
            self.title(f"Добавление записи в таблицу '{self.table_name_user}'")
            tk.Label(self, text="Наименование: ").grid(row=0, column=0, pady=5, padx=5)
            self.add_enty = tk.Entry(self, width=20)
            self.add_enty.grid(row=0, column=1, pady=5, padx=5)
            tk.Button(self, text="Отмена", width=20, command=self.quit_win).grid(row=1, column=0, pady=5, padx=5)
            tk.Button(self, text="Добавить", width=20, command=self.add).grid(row=1, column=1, pady=5, padx=5)
        
        elif operation == "delete":
            self.title(f"Удаление записи из таблицы '{self.table_name_user}'")
            tk.Label(self, text=f"Вы действиельно хотите удалить запись\nИз таблицы '{self.table_name_user}'?"
                     ).grid(row=0, column=0, columnspan=2, pady=5, padx=5)
            tk.Label(self, text=f"Значение: {self.value}").grid(row=1, column=0, 
                                                                                 columnspan=2, pady=5, padx=5)
            tk.Button(self, text="Да", command=self.delete, width=12).grid(row=2, column=0, pady=5, padx=5)
            tk.Button(self, text="Нет", command=self.quit_win, width=12).grid(row=2, column=1, pady=5, padx=5)
            
        elif operation == "change":
            self.title(f"Изменение записи в таблице '{self.table_name_user}'")
            tk.Label(self, text="текущее значение").grid(row=0, column=0, pady=5, padx=5)
            tk.Label(self, text="Новое значение").grid(row=0, column=1, pady=5, padx=5)

            tk.Label(self, text=f"{self.value}").grid(row=1, column=0, pady=5, padx=5)
            self.change_entry = tk.Entry(self, width=20)
            self.change_entry.grid(row=1, column=1, pady=5, padx=5)

            tk.Button(self, text="Отмена", width=20, command=self.quit_win).grid(row=2, column=0, pady=5, padx=5)
            tk.Button(self, text="Сохранить", width=20, command=self.change).grid(row=2, column=1, pady=5, padx=5)

    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()
    
    def add(self):
        new_value = self.add_enty.get()
        if new_value:
            try:
                conn = sqlite3.connect('resident.db')
                cursor = conn.cursor()
                cursor.execute(f"INSERT INTO {self.table_name_db} ({self.field_name}) VALUES (?)", (new_value,))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect('resident.db')
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM {self.table_name_db} WHERE {self.field_id} = ?", (self.id,))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))
    
    def change(self):
        new_value = self.change_entry.get()
        if new_value:
            try:
                conn = sqlite3.connect('resident.db')
                cursor = conn.cursor()
                cursor.execute(f"UPDATE {self.table_name_db} SET {self.field_name} = ? WHERE {self.field_id} = ?", 
                               (new_value, self.id))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
             showerror(title="Ошибка", message="Заполните все поля")

class WindowAnketa(tk.Toplevel):
    def __init__(self, operation: str, select_row = None):
        super().__init__()

        if select_row: self.select_row = select_row
        self.protocol('WM_DELETE_WINDOW', lambda: self.quit_win())

        conn = sqlite3.connect("resident.db")
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM gender")
        gender = []
        for item in cursor.fetchall(): gender.append(f"{item[0]}. {item[1]}")
        cursor.execute("SELECT * FROM citizen")
        citizen = []
        for item in cursor.fetchall(): citizen.append(f"{item[0]}. {item[1]}")
        conn.close

        if operation == "add":
            tk.Label(self, text="Полное имя гражданина: ").grid(row=0, column=0, pady=5, padx=5)
            self.id_citizen = ttk.Combobox(self, values=citizen)
            self.id_citizen.grid(row=0, column=1, pady=5, padx=5)

            tk.Label(self, text="Дата рождения: ").grid(row=1, column=0, pady=5, padx=5)
            self.date_of_birth = tk.Entry(self, width=20)
            self.date_of_birth.grid(row=1, column=1, pady=5, padx=5)

            tk.Label(self, text="Гражданство: ").grid(row=2, column=0, pady=5, padx=5)
            self.citizenship = tk.Entry(self, width=20)
            self.citizenship.grid(row=2, column=1, pady=5, padx=5)

            tk.Label(self, text="гендер: ").grid(row=3, column=0, pady=5, padx=5)
            self.id_gender = ttk.Combobox(self, values=gender)
            self.id_gender.grid(row=3, column=1, pady=5, padx=5)

            tk.Label(self, text="Адрес: ").grid(row=4, column=0, pady=5, padx=5)
            self.home_address = tk.Entry(self, width=20)
            self.home_address.grid(row=4, column=1, pady=5, padx=5)

            tk.Label(self, text="Место рождения: ").grid(row=5, column=0, pady=5, padx=5)
            self.place_of_birth = tk.Entry(self, width=20)
            self.place_of_birth.grid(row=5, column=1, pady=5, padx=5)

            tk.Label(self, text="ИНН: ").grid(row=6, column=0, pady=5, padx=5)
            self.inn = tk.Entry(self, width=20)
            self.inn.grid(row=6, column=1, pady=5, padx=5)

            tk.Label(self, text="Номер страховки: ").grid(row=7, column=0, pady=5, padx=5)
            self.insurance_number = tk.Entry(self, width=20)
            self.insurance_number.grid(row=7, column=1, pady=5, padx=5)

            tk.Label(self, text="Номер телефона: ").grid(row=8, column=0, pady=5, padx=5)
            self.phone_number = tk.Entry(self, width=20)
            self.phone_number.grid(row=8, column=1, pady=5, padx=5)

            tk.Label(self, text="Материальный статус: ").grid(row=9, column=0, pady=5, padx=5)
            self.marital_status = tk.Entry(self, width=20)
            self.marital_status.grid(row=9, column=1, pady=5, padx=5)

            tk.Label(self, text="Дополнительная информация: ").grid(row=10, column=0, pady=5, padx=5)
            self.additional_info = tk.Entry(self, width=20)
            self.additional_info.grid(row=10, column=1, pady=5, padx=5)

            tk.Label(self, text="Работодатель: ").grid(row=11, column=0, pady=5, padx=5)
            self.employer = tk.Entry(self, width=20)
            self.employer.grid(row=11, column=1, pady=5, padx=5)

            tk.Label(self, text="номер избирательного участка: ").grid(row=12, column=0, pady=5, padx=5)
            self.polling_station_number = tk.Entry(self, width=20)
            self.polling_station_number.grid(row=12, column=1, pady=5, padx=5)

            tk.Button(self, text="Отмена", width=20, command=self.quit_win).grid(row=13, column=0, pady=5, padx=5)
            tk.Button(self, text="Добавить", width=20, command=self.add).grid(row=13, column=1, pady=5, padx=5)
        
        elif operation == "delete":
            tk.Label(self, text=f"Вы действиельно хотите удалить запись?").grid(row=0, column=0, columnspan=2, pady=5, padx=5)
            tk.Label(self, text=f"Значение: {self.select_row[1]}", width=12).grid(row=1, column=0, columnspan=2, pady=5, padx=5)
            tk.Button(self, text="Да", command=self.delete).grid(row=2, column=0, pady=5, padx=5)
            tk.Button(self, text="Нет", command=self.quit_win).grid(row=2, column=1, pady=5, padx=5)
            
        elif operation == "change":
            tk.Label(self, text="текущее значение").grid(row=0, column=0, pady=5, padx=5)
            tk.Label(self, text="Новое значение").grid(row=0, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[1]).grid(row=1, column=0, pady=5, padx=5)
            self.id_citizen = ttk.Combobox(self, values=citizen)
            self.id_citizen.grid(row=1, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[2]).grid(row=2, column=0, pady=5, padx=5)
            self.date_of_birth = tk.Entry(self, width=20)
            self.date_of_birth.grid(row=2, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[3]).grid(row=3, column=0, pady=5, padx=5)
            self.citizenship = tk.Entry(self, width=20)
            self.citizenship.grid(row=3, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[4]).grid(row=4, column=0, pady=5, padx=5)
            self.id_gender = ttk.Combobox(self, values=gender)
            self.id_gender.grid(row=4, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[5]).grid(row=5, column=0, pady=5, padx=5)
            self.home_address = tk.Entry(self, width=20)
            self.home_address.grid(row=5, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[6]).grid(row=6, column=0, pady=5, padx=5)
            self.place_of_birth = tk.Entry(self, width=20)
            self.place_of_birth.grid(row=6, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[7]).grid(row=7, column=0, pady=5, padx=5)
            self.inn = tk.Entry(self, width=20)
            self.inn.grid(row=7, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[8]).grid(row=8, column=0, pady=5, padx=5)
            self.insurance_number = tk.Entry(self, width=20)
            self.insurance_number.grid(row=8, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[9]).grid(row=9, column=0, pady=5, padx=5)
            self.phone_number = tk.Entry(self, width=20)
            self.phone_number.grid(row=9, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[10]).grid(row=10, column=0, pady=5, padx=5)
            self.marital_status = tk.Entry(self, width=20)
            self.marital_status.grid(row=10, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[11]).grid(row=11, column=0, pady=5, padx=5)
            self.additional_info = tk.Entry(self, width=20)
            self.additional_info.grid(row=11, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[12]).grid(row=12, column=0, pady=5, padx=5)
            self.employer = tk.Entry(self, width=20)
            self.employer.grid(row=12, column=1, pady=5, padx=5)

            tk.Label(self, text=self.select_row[13]).grid(row=13, column=0, pady=5, padx=5)
            self.polling_station_number = tk.Entry(self, width=20)
            self.polling_station_number.grid(row=13, column=1, pady=5, padx=5)

            tk.Button(self, text="Отмена ", width=20, command=self.quit_win).grid(row=14, column=0, pady=5, padx=5)
            tk.Button(self, text="Сохранить", width=20, command=self.change).grid(row=14, column=1, pady=5, padx=5)

    def quit_win(self):
        win.deiconify()
        win.update_table()
        self.destroy()
    
    def add(self):
        id_citizen = self.id_citizen.get().split(".")[0]
        date_of_birth = self.date_of_birth.get()
        citizenship = self.citizenship.get()
        id_gender = self.id_gender.get().split(".")[0]
        home_address = self.home_address.get()
        place_of_birth = self.place_of_birth.get()
        inn = self.inn.get()
        insurance_number = self.insurance_number.get()
        phone_number = self.phone_number.get()
        marital_status = self.marital_status.get()
        additional_info = self.additional_info.get()
        employer = self.employer.get()
        polling_station_number = self.polling_station_number.get()
        
        if  (id_citizen and date_of_birth and citizenship and id_gender and home_address and place_of_birth and inn and 
             insurance_number and phone_number and marital_status and additional_info and employer and polling_station_number):
            try:
                conn = sqlite3.connect('resident.db')
                cursor = conn.cursor()
                cursor.execute(f"""INSERT INTO anketa (id_citizen, date_of_birth, citizenship, 
                               id_gender, home_address, place_of_birth, inn, insurance_number, phone_number, 
                               marital_status, additional_info, employer, polling_station_number) 
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", 
                               (id_citizen, date_of_birth, citizenship, id_gender, home_address, place_of_birth, 
                                inn, insurance_number, phone_number, marital_status, additional_info, employer, 
                                polling_station_number,))
                conn.commit()
                conn.close()
                self.quit_win()
            except sqlite3.Error as e:
                showerror(title="Ошибка", message=str(e))
        else:
            showerror(title="Ошибка", message="Заполните все поля")

    def delete(self):
        try:
            conn = sqlite3.connect('resident.db')
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM anketa WHERE id_anketa = ?", (self.select_row[0],))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))
    
    def change(self):
        id_citizen = self.id_citizen.get().split(".")[0]
        date_of_birth = self.date_of_birth.get() or self.select_row[2]
        citizenship = self.citizenship.get() or self.select_row[3]
        id_gender = self.id_gender.get().split(".")[0]
        home_address = self.home_address.get() or self.select_row[5]
        place_of_birth = self.place_of_birth.get() or self.select_row[6]
        inn = self.inn.get() or self.select_row[7]
        insurance_number = self.insurance_number.get() or self.select_row[8]
        phone_number = self.phone_number.get() or self.select_row[9]
        marital_status = self.marital_status.get() or self.select_row[10]
        additional_info = self.additional_info.get() or self.select_row[11]
        employer = self.employer.get() or self.select_row[12]
        polling_station_number = self.polling_station_number.get() or self.select_row[13]

        try:
            conn = sqlite3.connect('resident.db')
            cursor = conn.cursor()
            cursor.execute(f"""UPDATE anketa SET (id_citizen, date_of_birth, citizenship, id_gender, 
                            home_address, place_of_birth, inn, insurance_number, phone_number, marital_status, 
                           additional_info, employer, polling_station_number) = (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) 
                           WHERE id_anketa = {self.select_row[0]}""", 
                           (id_citizen, date_of_birth, citizenship, id_gender, home_address, place_of_birth, 
                            inn, insurance_number, phone_number, marital_status, additional_info, employer, 
                            polling_station_number,))
            conn.commit()
            conn.close()
            self.quit_win()
        except sqlite3.Error as e:
            showerror(title="Ошибка", message=str(e))

if __name__ == "__main__":
    db = DB()
    win = WindowMain()
    win.mainloop()