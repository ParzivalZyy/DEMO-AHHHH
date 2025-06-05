import sqlite3
import pandas as pd
import ttkbootstrap as tbs
from ttkbootstrap.constants import W, E, N, S
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import hashlib
import os
import re

class HotelManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.option_add("*Font", "TkDefaultFont 12")
        self.root.title("Система управления гостиницей")
        self.root.minsize(800, 600)
        self.root.configure(bg='#2f3542')
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 800
        window_height = 600
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        self.conn = sqlite3.connect('hotel.db', detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
        self.cursor = self.conn.cursor()
        self.current_user = None
        self.init_db()
        self.load_rooms_from_excel('Номерной фонд.xlsx')
        self.create_login_form()

    def init_db(self):
        self.cursor.executescript("""
            CREATE TABLE IF NOT EXISTS guests (
                guestID INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                phone TEXT NOT NULL UNIQUE,
                email TEXT NOT NULL UNIQUE,
                passport TEXT NOT NULL UNIQUE,
                preferences TEXT
            );

            CREATE TABLE IF NOT EXISTS guest_requests (
                requestID INTEGER PRIMARY KEY AUTOINCREMENT,
                guest_id INTEGER,
                request TEXT,
                status TEXT CHECK(status IN ('Новая', 'В работе', 'Выполнено')),
                request_date DATE DEFAULT CURRENT_DATE,
                FOREIGN KEY(guest_id) REFERENCES guests(guestID)
            );

            CREATE TABLE IF NOT EXISTS rooms (
                roomID INTEGER PRIMARY KEY AUTOINCREMENT,
                room_number TEXT NOT NULL UNIQUE,
                status TEXT CHECK(status IN ('Свободен', 'Занят', 'Грязный', 'Назначен к уборке', 'Чистый')) DEFAULT 'Свободен',
                price_per_night REAL NOT NULL,
                floor INTEGER
            );

            CREATE TABLE IF NOT EXISTS bookings (
                bookingID INTEGER PRIMARY KEY AUTOINCREMENT,
                guest_id INTEGER,
                room_id INTEGER,
                check_in DATE,
                check_out DATE,
                booking_date DATE DEFAULT CURRENT_DATE,
                status TEXT CHECK(status IN ('Забронировано', 'Заселен', 'Отменено', 'Завершено')),
                FOREIGN KEY(guest_id) REFERENCES guests(guestID),
                FOREIGN KEY(room_id) REFERENCES rooms(roomID)
            );

            CREATE TABLE IF NOT EXISTS payments (
                paymentID INTEGER PRIMARY KEY AUTOINCREMENT,
                booking_id INTEGER,
                payment_date DATE DEFAULT CURRENT_DATE,
                amount REAL,
                receipt_number TEXT,
                FOREIGN KEY(booking_id) REFERENCES bookings(bookingID)
            );

            CREATE TABLE IF NOT EXISTS staff (
                staffID INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT,
                role TEXT CHECK(role IN ('Администратор', 'Руководитель', 'Уборщик')),
                login TEXT UNIQUE,
                password TEXT,
                last_login DATE,
                login_attempts INTEGER DEFAULT 0,
                is_blocked INTEGER DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS cleaning (
                cleaningID INTEGER PRIMARY KEY AUTOINCREMENT,
                room_id INTEGER,
                staff_id INTEGER,
                scheduled_date DATE,
                status TEXT CHECK(status IN ('Назначено', 'Выполнено')),
                FOREIGN KEY(room_id) REFERENCES rooms(roomID),
                FOREIGN KEY(staff_id) REFERENCES staff(staffID)
            );
        """)
        self.cursor.execute("INSERT OR IGNORE INTO staff (full_name, role, login, password) VALUES (?, ?, ?, ?)",
                           ("Админ", "Администратор", "AAA", self.hash_password("121212")))
        self.conn.commit()

    def load_rooms_from_excel(self, filename):
        if not os.path.exists(filename):
            print(f"Файл {filename} не найден.")
            return
        df = pd.read_excel(filename)
        for _, row in df.iterrows():
            self.cursor.execute("""
                INSERT OR IGNORE INTO rooms (room_number, floor, price_per_night, status)
                VALUES (?, ?, ?, 'Свободен')
            """, (str(row['Номер']), row['Этаж'], self.get_price(row['Категория']),))
        self.conn.commit()

    def get_price(self, category):
        prices = {
            'Одноместный стандарт': 1000,
            'Одноместный эконом': 800,
            'Стандарт двухместный с 2 раздельными кроватями': 1500,
            'Эконом двухместный с 2 раздельными кроватями': 1200,
            '3-местный бюджет': 1800,
            'Бизнес с 1 или 2 кроватями': 2000,
            'Двухкомнатный двухместный стандарт с 1 или 2 кроватями': 2200,
            'Студия': 2500,
            'Люкс с 2 двуспальными кроватями': 3000
        }
        return prices.get(category, 1000)

    def hash_password(self, password):
        return hashlib.sha256(password.encode()).hexdigest()

    def create_login_form(self):
        self.clear_frame()
        self.root.configure(bg='#2f3542')
        main_frame = tbs.Frame(self.root, bootstyle="primary")
        main_frame.pack(expand=True)
        
        content_frame = tbs.Frame(main_frame, bootstyle="primary", padding=20)
        content_frame.pack(expand=True)
        
        frame = tbs.Frame(content_frame, bootstyle="primary")
        frame.pack()
        
        login_frame = tbs.Frame(frame, bootstyle="primary")
        login_frame.grid(row=0, column=0, columnspan=2, pady=10)
        
        tbs.Label(login_frame, text="Логин:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.login_var = tk.StringVar()
        tbs.Entry(login_frame, textvariable=self.login_var, bootstyle="primary").grid(row=0, column=1, sticky=(W, E), padx=5, pady=5)
        
        tbs.Label(login_frame, text="Пароль:", bootstyle="inverse-primary").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        self.password_var = tk.StringVar()
        tbs.Entry(login_frame, textvariable=self.password_var, show="*", bootstyle="primary").grid(row=1, column=1, sticky=(W, E), padx=5, pady=5)
        
        button_frame = tbs.Frame(frame, bootstyle="primary")
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)
        button_frame.grid_columnconfigure(0, weight=1)
        
        tbs.Button(button_frame, text="Войти", command=self.authenticate, bootstyle="primary-outline").grid(row=0, column=0, pady=10)

    def authenticate(self):
        login = self.login_var.get()
        password = self.hash_password(self.password_var.get())
        self.cursor.execute("SELECT * FROM staff WHERE login = ?", (login,))
        user = self.cursor.fetchone()
        if not user:
            messagebox.showerror("Ошибка", "Несуществующий логин или пароль. Пожалуйста, проверьте введенные данные.")
            return
        if user[2] == 'Администратор':
            if user[4] == password:
                now = datetime.now().date()
                self.cursor.execute("UPDATE staff SET last_login = ? WHERE staffID = ?", (now, user[0]))
                self.conn.commit()
                self.current_user = user
                self.create_main_menu()
            else:
                messagebox.showerror("Ошибка", "Вы ввели неверный логин или пароль. Пожалуйста, проверьте введенные данные или обратитесь к админу.")
            return

        if user[7] == 1:
            messagebox.showerror("Ошибка", "Вы заблокированы. Обратитесь к администратору.")
            return
        if user[4] == password:
            now = datetime.now().date()
            self.cursor.execute("UPDATE staff SET login_attempts = 0, last_login = ? WHERE staffID = ?", (now, user[0]))
            self.conn.commit()
            self.current_user = user
            self.create_main_menu()
        else:
            attempts = user[6] + 1
            if attempts >= 3:
                self.cursor.execute("UPDATE staff SET is_blocked = 1, login_attempts = 0 WHERE staffID = ?", (user[0],))
                messagebox.showerror("Ошибка", "Вы заблокированы. Обратитесь к администратору.")
            else:
                self.cursor.execute("UPDATE staff SET login_attempts = ? WHERE staffID = ?", (attempts, user[0]))
                messagebox.showerror("Ошибка", "Вы ввели неверный логин или пароль. Пожалуйста, проверьте введенные данные или обратитесь к админу.")
            self.conn.commit()

    def change_password(self):
        current = self.hash_password(self.current_password.get())
        new = self.new_password.get()
        confirm = self.confirm_password.get()
        self.cursor.execute("SELECT password FROM staff WHERE staffID = ?", (self.current_user[0],))
        if current != self.cursor.fetchone()[0]:
            messagebox.showerror("Ошибка", "Неверный текущий пароль")
            return
        if new != confirm:
            messagebox.showerror("Ошибка", "Новый пароль и подтверждение не совпадают")
            return
        if not new:
            messagebox.showerror("Ошибка", "Пароль не может быть пустым")
            return
        self.cursor.execute("UPDATE staff SET password = ? WHERE staffID = ?", (self.hash_password(new), self.current_user[0]))
        self.conn.commit()
        self.create_main_menu()

    def create_base_form(self, back_command):
        self.clear_frame()
        main_frame = tbs.Frame(self.root, bootstyle="primary")
        main_frame.pack(expand=True, fill='both')
        back_button = tbs.Button(main_frame, text="Назад", command=back_command, bootstyle="primary-outline", width=10)
        back_button.pack(anchor='nw', padx=10, pady=10)
        content_frame = tbs.Frame(main_frame, bootstyle="primary", padding=20)
        content_frame.pack(expand=True)
        return content_frame

    def create_main_menu(self):
        self.clear_frame()
        self.root.configure(bg='#2f3542')
        main_frame = tbs.Frame(self.root, bootstyle="primary")
        main_frame.pack(expand=True, fill='both')

        content_frame = tbs.Frame(main_frame, bootstyle="primary", padding=20)
        content_frame.pack(expand=True)

        frame = tbs.Frame(content_frame, bootstyle="primary")
        frame.pack(expand=True)

        button_frame = tbs.Frame(frame, bootstyle="primary")
        button_frame.pack(expand=True)

        if self.current_user[2] == 'Администратор':
            tbs.Button(button_frame, text="Добавить пользователя", command=self.create_add_user_form, bootstyle="primary-outline", width=40).pack(pady=15)
            tbs.Button(button_frame, text="Управление бронированиями", command=self.create_booking_form, bootstyle="primary-outline", width=40).pack(pady=15)
            tbs.Button(button_frame, text="Управление номерами", command=self.create_room_management_form, bootstyle="primary-outline", width=40).pack(pady=15)
            tbs.Button(button_frame, text="График уборки", command=self.create_cleaning_schedule_form, bootstyle="primary-outline", width=40).pack(pady=15)
            tbs.Button(button_frame, text="Отчеты", command=self.create_reports_form, bootstyle="primary-outline", width=40).pack(pady=15)
            tbs.Button(button_frame, text="Разблокировать пользователей", command=self.create_unblock_users_form, bootstyle="danger-outline", width=40).pack(pady=15)
        elif self.current_user[2] == 'Уборщик':
            tbs.Button(button_frame, text="Управление номерами", command=self.create_room_management_form, bootstyle="primary-outline", width=40).pack(pady=15)
            tbs.Button(button_frame, text="График уборки", command=self.create_cleaning_schedule_form, bootstyle="primary-outline", width=40).pack(pady=15)
        elif self.current_user[2] == 'Руководитель':
            tbs.Button(button_frame, text="Управление бронированиями", command=self.create_booking_form, bootstyle="primary-outline", width=40).pack(pady=15)
            tbs.Button(button_frame, text="Управление номерами", command=self.create_room_management_form, bootstyle="primary-outline", width=40).pack(pady=15)
            tbs.Button(button_frame, text="График уборки", command=self.create_cleaning_schedule_form, bootstyle="primary-outline", width=40).pack(pady=15)
            tbs.Button(button_frame, text="Отчеты", command=self.create_reports_form, bootstyle="primary-outline", width=40).pack(pady=15)

    def create_add_user_form(self):
        content_frame = self.create_base_form(self.create_main_menu)
        form_frame = tbs.Frame(content_frame, bootstyle="primary")
        form_frame.pack(expand=True)
        tbs.Label(form_frame, text="Имя:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.full_name_var = tk.StringVar()
        tbs.Entry(form_frame, textvariable=self.full_name_var, bootstyle="primary").grid(row=0, column=1, sticky=(W, E), padx=5, pady=5)
        tbs.Label(form_frame, text="Логин:", bootstyle="inverse-primary").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        self.new_login_var = tk.StringVar()
        tbs.Entry(form_frame, textvariable=self.new_login_var, bootstyle="primary").grid(row=1, column=1, sticky=(W, E), padx=5, pady=5)
        tbs.Label(form_frame, text="Пароль:", bootstyle="inverse-primary").grid(row=2, column=0, sticky=W, padx=5, pady=5)
        self.new_password_var = tk.StringVar()
        tbs.Entry(form_frame, textvariable=self.new_password_var, bootstyle="primary").grid(row=2, column=1, sticky=(W, E), padx=5, pady=5)
        tbs.Label(form_frame, text="Роль:", bootstyle="inverse-primary").grid(row=3, column=0, sticky=W, padx=5, pady=5)
        self.role_var = tk.StringVar()
        tbs.Combobox(form_frame, textvariable=self.role_var, values=['Администратор', 'Руководитель', 'Уборщик'], bootstyle="primary").grid(row=3, column=1, sticky=(W, E), padx=5, pady=5)
        button_frame = tbs.Frame(form_frame, bootstyle="primary")
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)
        button_frame.grid_columnconfigure(0, weight=1)
        tbs.Button(button_frame, text="Добавить", command=self.add_user, bootstyle="primary").pack(pady=10)

    def add_user(self):
        login = self.new_login_var.get()
        password = self.hash_password(self.new_password_var.get())
        self.cursor.execute("SELECT login FROM staff WHERE login = ?", (login,))
        if self.cursor.fetchone():
            messagebox.showerror("Ошибка", "Пользователь с таким логином уже существует")
            return
        self.cursor.execute("INSERT INTO staff (full_name, role, login, password) VALUES (?, ?, ?, ?)",
                        (self.full_name_var.get(), self.role_var.get(), login, password))
        self.conn.commit()
        messagebox.showinfo("Успех", "Пользователь успешно добавлен")
        self.create_main_menu()

    def create_booking_form(self):
        content_frame = self.create_base_form(self.create_main_menu)
        form_frame = tbs.Frame(content_frame, bootstyle="primary")
        form_frame.pack(expand=True)
        
        guest_name_var = tk.StringVar()
        phone_var = tk.StringVar()
        email_var = tk.StringVar()
        passport_var = tk.StringVar()
        room_selection_var = tk.StringVar()
        check_in_var = tk.StringVar()
        check_out_var = tk.StringVar()

        tbs.Label(form_frame, text="Имя гостя:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        tbs.Entry(form_frame, textvariable=guest_name_var, bootstyle="primary").grid(row=0, column=1, sticky=(W, E), padx=5, pady=5)
        
        tbs.Label(form_frame, text="Номер телефона:", bootstyle="inverse-primary").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        tbs.Entry(form_frame, textvariable=phone_var, bootstyle="primary").grid(row=1, column=1, sticky=(W, E), padx=5, pady=5)
        
        tbs.Label(form_frame, text="Email:", bootstyle="inverse-primary").grid(row=2, column=0, sticky=W, padx=5, pady=5)
        tbs.Entry(form_frame, textvariable=email_var, bootstyle="primary").grid(row=2, column=1, sticky=(W, E), padx=5, pady=5)
        
        tbs.Label(form_frame, text="Паспорт:", bootstyle="inverse-primary").grid(row=3, column=0, sticky=W, padx=5, pady=5)
        tbs.Entry(form_frame, textvariable=passport_var, bootstyle="primary").grid(row=3, column=1, sticky=(W, E), padx=5, pady=5)
        
        tbs.Label(form_frame, text="Номер комнаты:", bootstyle="inverse-primary").grid(row=4, column=0, sticky=W, padx=5, pady=5)
        rooms_data = self.cursor.execute("SELECT roomID, room_number, floor FROM rooms WHERE status IN ('Свободен', 'Чистый')").fetchall()
        
        room_categories_excel = {}
        if os.path.exists('Номерной фонд.xlsx'):
            df = pd.read_excel('Номерной фонд.xlsx')
            room_categories_excel = dict(zip(df['Номер'].astype(str), df['Категория']))
            
        booking_room_map = {f"{r[1]} ({room_categories_excel.get(str(r[1]), 'Не указана')}, этаж {r[2]})": r[0] for r in rooms_data}
        
        room_cb = tbs.Combobox(form_frame, textvariable=room_selection_var, values=list(booking_room_map.keys()), bootstyle="primary", state="readonly")
        room_cb.grid(row=4, column=1, sticky=(W, E), padx=5, pady=5)
        
        tbs.Label(form_frame, text="Дата заезда:", bootstyle="inverse-primary").grid(row=5, column=0, sticky=W, padx=5, pady=5)
        tbs.Entry(form_frame, textvariable=check_in_var, bootstyle="primary").grid(row=5, column=1, sticky=(W, E), padx=5, pady=5)
        
        tbs.Label(form_frame, text="Дата выезда:", bootstyle="inverse-primary").grid(row=6, column=0, sticky=W, padx=5, pady=5)
        tbs.Entry(form_frame, textvariable=check_out_var, bootstyle="primary").grid(row=6, column=1, sticky=(W, E), padx=5, pady=5)

        def is_valid_phone(phone):
            return bool(re.fullmatch(r"(\+7|8)\d{10}$", phone))

        def is_valid_email(email):
            return bool(re.fullmatch(r"[^@]+@(?:inbox|mail|gmail)\.[a-zA-Z]{2,}$", email))

        def register_guest_nested():
            phone = phone_var.get()
            email = email_var.get()
            if not is_valid_phone(phone):
                messagebox.showerror("Ошибка", "Введите корректный российский номер телефона (+7XXXXXXXXXX или 8XXXXXXXXXX)")
                raise Exception("Invalid phone")
            if not is_valid_email(email):
                messagebox.showerror("Ошибка", "Введите корректный email с доменом @inbox, @mail или @gmail")
                raise Exception("Invalid email")
            self.cursor.execute("INSERT OR IGNORE INTO guests (full_name, phone, email, passport) VALUES (?, ?, ?, ?)",
                              (guest_name_var.get(), phone, email, passport_var.get()))
            self.conn.commit()
            self.cursor.execute("SELECT guestID FROM guests WHERE passport = ?", (passport_var.get(),))
            guest_fetch = self.cursor.fetchone()
            if guest_fetch:
                return guest_fetch[0]
            else:
                messagebox.showerror("Ошибка регистрации", "Не удалось зарегистрировать или найти гостя.")
                raise Exception("Guest registration failed")


        def create_booking_action():
            guest_id = register_guest_nested()
            
            room_key_selected = room_selection_var.get()
            if not room_key_selected:
                messagebox.showerror("Ошибка", "Выберите номер комнаты")
                return
            
            room_id_selected = booking_room_map.get(room_key_selected) 
            if not room_id_selected:
                messagebox.showerror("Ошибка", "Выбранная комната не найдена в системе.")
                return
                
            check_in_date_str = check_in_var.get()
            check_out_date_str = check_out_var.get()

            if not (check_in_date_str and check_out_date_str):
                messagebox.showerror("Ошибка", "Даты заезда и выезда должны быть указаны.")
                return

            check_in = datetime.strptime(check_in_date_str, '%Y-%m-%d').date()
            check_out = datetime.strptime(check_out_date_str, '%Y-%m-%d').date()

            if check_in >= check_out:
                messagebox.showerror("Ошибка", "Дата заезда должна быть раньше даты выезда")
                return

            conflict = self.cursor.execute("""
                SELECT 1 FROM bookings 
                WHERE room_id = ? 
                AND status IN ('Забронировано', 'Заселен')
                AND NOT (check_out <= ? OR check_in >= ?)
            """, (room_id_selected, check_in, check_out)).fetchone()
            
            if conflict:
                messagebox.showerror("Ошибка", "Комната занята на указанные даты")
                return
            
            self.cursor.execute("""
                INSERT INTO bookings (guest_id, room_id, check_in, check_out, status)
                VALUES (?, ?, ?, ?, 'Забронировано')
            """, (guest_id, room_id_selected, check_in, check_out))
            self.cursor.execute("UPDATE rooms SET status = 'Занят' WHERE roomID = ?", (room_id_selected,))
            self.conn.commit()
            messagebox.showinfo("Успех", "Бронирование успешно создано")
            self.create_main_menu()
        
        button_frame = tbs.Frame(form_frame, bootstyle="primary")
        button_frame.grid(row=7, column=0, columnspan=2, pady=10)
        button_frame.grid_columnconfigure(0, weight=1)
        tbs.Button(button_frame, text="Забронировать", command=create_booking_action, bootstyle="primary").pack(pady=10)

    def create_room_management_form(self):
        content_frame = self.create_base_form(self.create_main_menu)
        table_frame = tbs.Frame(content_frame, bootstyle="primary")
        table_frame.pack(expand=True, fill='both', padx=20, pady=(0, 0))
        tree = tbs.Treeview(
            table_frame,
            columns=('Номер', 'Этаж', 'Категория', 'Статус', 'Цена'),
            show='headings',
            bootstyle="primary"
        )
        tree.heading('Номер', text='Номер')
        tree.heading('Этаж', text='Этаж')
        tree.heading('Категория', text='Категория')
        tree.heading('Статус', text='Статус')
        tree.heading('Цена', text='Цена за ночь')
        tree.column('Номер', anchor='center', width=100)
        tree.column('Этаж', anchor='center', width=80)
        tree.column('Категория', anchor='center', width=200)
        tree.column('Статус', anchor='center', width=120)
        tree.column('Цена', anchor='center', width=100)
        tree.pack(expand=True, fill='both', pady=(0, 0), padx=0, side='top')
        
        room_categories_excel = {}
        if os.path.exists('Номерной фонд.xlsx'):
            df = pd.read_excel('Номерной фонд.xlsx')
            room_categories_excel = dict(zip(df['Номер'].astype(str), df['Категория']))

        rooms = self.cursor.execute("SELECT room_number, floor, status, price_per_night FROM rooms").fetchall()
        for room in rooms:
            category = room_categories_excel.get(str(room[0]), "Не указана")
            tree.insert('', 'end', values=(room[0], room[1], category, room[2], room[3]))
        table_frame.pack_configure(pady=(10, 10))

    def create_cleaning_schedule_form(self):
        content_frame = self.create_base_form(self.create_main_menu)
        tree = tbs.Treeview(content_frame, columns=('Номер', 'Дата', 'Статус'), show='headings', bootstyle="primary")
        tree.heading('Номер', text='Номер')
        tree.heading('Дата', text='Дата')
        tree.heading('Статус', text='Статус')
        tree.column('Номер', anchor='center')
        tree.column('Дата', anchor='center')
        tree.column('Статус', anchor='center')
        tree.pack(expand=True, fill='both', pady=10)
        cleanings = self.cursor.execute("""
            SELECT r.room_number, c.scheduled_date, c.status 
            FROM cleaning c 
            JOIN rooms r ON c.room_id = r.roomID 
            WHERE c.status = 'Назначено' 
            AND (? = 'Администратор' OR ? = 'Руководитель' OR c.staff_id = ?)
        """, (self.current_user[2], self.current_user[2], self.current_user[0])).fetchall()
        
        for cleaning_item in cleanings:
            tree.insert('', 'end', values=cleaning_item)
            
        button_frame = tbs.Frame(content_frame, bootstyle="primary")
        button_frame.pack(pady=10)

        if self.current_user[2] in ['Администратор', 'Руководитель']:
             tbs.Button(button_frame, text="Запланировать уборку", command=self.open_plan_cleaning_window, bootstyle="primary-outline").pack(side='left', padx=(0, 10))
        
        if self.current_user[2] == 'Уборщик' or self.current_user[2] in ['Администратор', 'Руководитель']: 
            tbs.Button(button_frame, text="Завершить уборку", command=lambda: self.complete_cleaning(tree), bootstyle="primary-outline").pack(side='left')


    def open_plan_cleaning_window(self):
        plan_win = tk.Toplevel(self.root)
        plan_win.title("Запланировать уборку")
        plan_win.geometry("400x220")
        plan_win.grab_set()
        plan_win.resizable(False, False)
        frame = tbs.Frame(plan_win, bootstyle="primary", padding=20)
        frame.pack(expand=True, fill='both')

        tbs.Label(frame, text="Номер комнаты:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        cleaning_room_var = tk.StringVar()
        rooms_for_cleaning = self.cursor.execute("""
            SELECT roomID, room_number, floor FROM rooms 
            WHERE status IN ('Грязный', 'Назначен к уборке', 'Занят')
            AND roomID NOT IN (
                SELECT room_id 
                FROM cleaning 
                WHERE status = 'Назначено'
            )
        """).fetchall()
        room_categories_excel = {}
        if os.path.exists('Номерной фонд.xlsx'):
            df = pd.read_excel('Номерной фонд.xlsx')
            room_categories_excel = dict(zip(df['Номер'].astype(str), df['Категория']))
        cleaning_room_map = {f"{r[1]} ({room_categories_excel.get(str(r[1]), 'Не указана')}, этаж {r[2]})": r[0] for r in rooms_for_cleaning}
        room_cb = tbs.Combobox(frame, textvariable=cleaning_room_var, values=list(cleaning_room_map.keys()), bootstyle="primary", state="readonly")
        room_cb.grid(row=0, column=1, sticky=(W, E), padx=5, pady=5)

        tbs.Label(frame, text="Сотрудник:", bootstyle="inverse-primary").grid(row=1, column=0, sticky=W, padx=5, pady=5)
        cleaning_staff_var = tk.StringVar()
        staff_data = self.cursor.execute("SELECT staffID, full_name FROM staff WHERE role = 'Уборщик' AND is_blocked = 0").fetchall()
        cleaning_staff_map = {f"{s[1]} (ID:{s[0]})": s[0] for s in staff_data}
        staff_cb = tbs.Combobox(frame, textvariable=cleaning_staff_var, values=list(cleaning_staff_map.keys()), bootstyle="primary", state="readonly")
        staff_cb.grid(row=1, column=1, sticky=(W, E), padx=5, pady=5)

        tbs.Label(frame, text="Дата:", bootstyle="inverse-primary").grid(row=2, column=0, sticky=W, padx=5, pady=5)
        cleaning_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tbs.Entry(frame, textvariable=cleaning_date_var, bootstyle="primary").grid(row=2, column=1, sticky=(W, E), padx=5, pady=5)

        def plan_cleaning_action():
            room_key = cleaning_room_var.get()
            staff_key = cleaning_staff_var.get()
            scheduled_date_str = cleaning_date_var.get()
            if not room_key or not staff_key or not scheduled_date_str:
                messagebox.showerror("Ошибка", "Заполните все поля", parent=plan_win)
                return
            room_id = cleaning_room_map.get(room_key)
            staff_id = cleaning_staff_map.get(staff_key)
            if not room_id:
                messagebox.showerror("Ошибка", "Выбранная комната не найдена.", parent=plan_win)
                return
            if not staff_id:
                messagebox.showerror("Ошибка", "Выбранный сотрудник не найден.", parent=plan_win)
                return
            scheduled_date = datetime.strptime(scheduled_date_str, '%Y-%m-%d').date()
            exists = self.cursor.execute(
                "SELECT 1 FROM cleaning WHERE room_id = ? AND scheduled_date = ? AND status = 'Назначено'",
                (room_id, scheduled_date)
            ).fetchone()
            if exists:
                messagebox.showerror("Ошибка", "Уборка для этой комнаты на эту дату уже запланирована", parent=plan_win)
                return
            self.cursor.execute(
                "INSERT INTO cleaning (room_id, staff_id, scheduled_date, status) VALUES (?, ?, ?, 'Назначено')",
                (room_id, staff_id, scheduled_date)
            )
            self.cursor.execute("UPDATE rooms SET status = 'Назначен к уборке' WHERE roomID = ?", (room_id,))
            self.conn.commit()
            messagebox.showinfo("Успех", "Уборка успешно запланирована", parent=plan_win)
            plan_win.destroy()
            self.create_cleaning_schedule_form()
        tbs.Button(frame, text="Запланировать", command=plan_cleaning_action, bootstyle="primary").grid(row=3, column=0, columnspan=2, pady=15)

    def complete_cleaning(self, tree):
        selected_item_id = tree.selection()
        if not selected_item_id:
            messagebox.showerror("Ошибка", "Выберите уборку для завершения")
            return
        
        selected_values = tree.item(selected_item_id[0])['values']
        room_number = selected_values[0]
        scheduled_date_str = selected_values[1] 
        
        room_info = self.cursor.execute("SELECT roomID FROM rooms WHERE room_number = ?", (room_number,)).fetchone()
        if not room_info:
            messagebox.showerror("Ошибка", "Комната не найдена.")
            return
        room_id = room_info[0]
        
        cleaning_task_query = """
            UPDATE cleaning SET status = 'Выполнено' 
            WHERE room_id = ? AND scheduled_date = ? AND status = 'Назначено'
        """
        params = [room_id, scheduled_date_str]

        if self.current_user[2] == 'Уборщик':
            cleaning_task_query += " AND staff_id = ?"
            params.append(self.current_user[0])
        
        self.cursor.execute(cleaning_task_query, tuple(params))
        
        if self.cursor.rowcount > 0:
            self.cursor.execute("UPDATE rooms SET status = 'Чистый' WHERE roomID = ?", (room_id,))
            self.conn.commit()
            messagebox.showinfo("Успех", "Уборка завершена")
        else:
            messagebox.showerror("Ошибка", "Не удалось обновить статус уборки. Возможно, она уже выполнена или не найдена.")
            
        self.create_cleaning_schedule_form()


    def create_reports_form(self):
        content_frame = self.create_base_form(self.create_main_menu)
        form_frame = tbs.Frame(content_frame, bootstyle="primary")
        form_frame.pack(expand=True)
        tbs.Label(form_frame, text="Дата отчета:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, padx=5, pady=5)
        self.report_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tbs.Entry(form_frame, textvariable=self.report_date_var, bootstyle="primary").grid(row=0, column=1, sticky=(W, E), padx=5, pady=5)
        button_frame = tbs.Frame(form_frame, bootstyle="primary")
        button_frame.grid(row=1, column=0, columnspan=2, pady=10)
        button_frame.grid_columnconfigure(0, weight=1)
        tbs.Button(button_frame, text="Сформировать отчет", command=self.generate_report, bootstyle="primary-outline").pack(pady=10)

    def generate_report(self):
        try:
            report_date_str = self.report_date_var.get()
            report_date = datetime.strptime(report_date_str, '%Y-%m-%d').date()
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ГГГГ-ММ-ДД.")
            return

        total_rooms = self.cursor.execute("SELECT COUNT(*) FROM rooms").fetchone()[0]
        occupied_rooms_count = self.cursor.execute("""
            SELECT COUNT(DISTINCT room_id) FROM bookings 
            WHERE status IN ('Забронировано', 'Заселен') 
            AND date(check_in) <= date(?) AND date(check_out) > date(?)
        """, (report_date, report_date)).fetchone()[0]
        occupancy_rate = (occupied_rooms_count / total_rooms) * 100 if total_rooms > 0 else 0
        revenue_data = self.cursor.execute("SELECT SUM(amount) FROM payments WHERE date(payment_date) = date(?)", (report_date,)).fetchone()
        daily_revenue = revenue_data[0] if revenue_data and revenue_data[0] is not None else 0
        adr = daily_revenue / occupied_rooms_count if occupied_rooms_count > 0 else 0
        revpar = daily_revenue / total_rooms if total_rooms > 0 else 0
        messagebox.showinfo("Отчет", f"Дата: {report_date_str}\n"
                                   f"Всего номеров: {total_rooms}\n"
                                   f"Занято номеров: {occupied_rooms_count}\n"
                                   f"Процент загрузки: {occupancy_rate:.2f}%\n"
                                   f"Доход за день: {daily_revenue:.2f}\n"
                                   f"ADR (по доходу дня): {adr:.2f}\n"
                                   f"RevPAR (по доходу дня): {revpar:.2f}")


    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def create_unblock_users_form(self):
        content_frame = self.create_base_form(self.create_main_menu)
        form_frame = tbs.Frame(content_frame, bootstyle="primary")
        form_frame.pack(expand=True)
        
        blocked_users = self.cursor.execute("""
            SELECT staffID, full_name, login, role 
            FROM staff 
            WHERE is_blocked = 1
        """).fetchall()
        
        if not blocked_users:
            tbs.Label(form_frame, text="Нет заблокированных пользователей", 
                     bootstyle="inverse-primary").pack(pady=20)
            return
            
        tree = tbs.Treeview(form_frame, columns=('ID', 'Имя', 'Логин', 'Роль'), 
                           show='headings', bootstyle="primary")
        tree.heading('ID', text='ID')
        tree.heading('Имя', text='Имя')
        tree.heading('Логин', text='Логин')
        tree.heading('Роль', text='Роль')
        
        for user in blocked_users:
            tree.insert('', 'end', values=user)
        
        tree.pack(pady=20, fill='both', expand=True)
        
        def unblock_selected():
            selection = tree.selection()
            if not selection:
                messagebox.showerror("Ошибка", "Выберите пользователя для разблокировки")
                return
                
            selected_id = tree.item(selection[0])['values'][0]
            self.cursor.execute("""
                UPDATE staff 
                SET is_blocked = 0, login_attempts = 0 
                WHERE staffID = ?
            """, (selected_id,))
            self.conn.commit()
            messagebox.showinfo("Успех", "Пользователь разблокирован")
            self.create_unblock_users_form()  
            
        tbs.Button(form_frame, text="Разблокировать выбранного пользователя", 
                  command=unblock_selected, 
                  bootstyle="danger-outline").pack(pady=10)

    def __del__(self):
        if hasattr(self, 'conn') and self.conn:
            self.conn.close()

if __name__ == "__main__":
    root = tbs.Window(themename="darkly")
    app = HotelManagementApp(root)
    root.mainloop()

