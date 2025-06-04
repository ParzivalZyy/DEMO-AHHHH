import sqlite3
import pandas as pd
import ttkbootstrap as tbs
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import messagebox
from datetime import datetime, timedelta
import hashlib
import os

class HotelManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Система управления гостиницей")
        self.root.minsize(800, 600)
        self.conn = sqlite3.connect('hotel.db')
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
        frame = tbs.Frame(self.root, bootstyle="primary", padding=10)
        frame.pack(expand=True)
        tbs.Label(frame, text="Логин:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, pady=5)
        self.login_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.login_var, bootstyle="primary").grid(row=0, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Пароль:", bootstyle="inverse-primary").grid(row=1, column=0, sticky=W, pady=5)
        self.password_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.password_var, show="*", bootstyle="primary").grid(row=1, column=1, sticky=(W, E), pady=5)
        tbs.Button(frame, text="Войти", command=self.authenticate, bootstyle="primary").grid(row=2, column=0, columnspan=2, pady=10)

    def authenticate(self):
        login = self.login_var.get()
        password = self.hash_password(self.password_var.get())
        self.cursor.execute("SELECT * FROM staff WHERE login = ? AND is_blocked = 0", (login,))
        user = self.cursor.fetchone()
        if not user:
            self.cursor.execute("SELECT * FROM staff WHERE login = ?", (login,))
            user = self.cursor.fetchone()
            if user and user[7] == 1:
                messagebox.showerror("Ошибка", "Вы заблокированы. Обратитесь к администратору.")
                return
            messagebox.showerror("Ошибка", "Вы ввели неверный логин или пароль. Пожалуйста, проверьте введенные данные.")
            return
        if user[4] == password:
            last_login = user[5]
            if last_login and (datetime.now().date() - datetime.strptime(last_login, '%Y-%m-%d').date()).days > 30:
                self.cursor.execute("UPDATE staff SET is_blocked = 1 WHERE staffID = ?", (user[0],))
                self.conn.commit()
                messagebox.showerror("Ошибка", "Вы заблокированы из-за неактивности. Обратитесь к администратору.")
                return
            self.cursor.execute("UPDATE staff SET login_attempts = 0, last_login = ? WHERE staffID = ?",
                              (datetime.now().date(), user[0]))
            self.conn.commit()
            self.current_user = user
            if user[4] == self.hash_password("default"):
                self.create_change_password_form()
            else:
                messagebox.showinfo("Успех", "Вы успешно авторизовались")
                self.create_main_menu()
        else:
            attempts = user[6] + 1
            if attempts >= 3:
                self.cursor.execute("UPDATE staff SET is_blocked = 1, login_attempts = 0 WHERE staffID = ?", (user[0],))
                messagebox.showerror("Ошибка", "Вы заблокированы. Обратитесь к администратору.")
            else:
                self.cursor.execute("UPDATE staff SET login_attempts = ? WHERE staffID = ?", (attempts, user[0]))
                messagebox.showerror("Ошибка", "Вы ввели неверный логин или пароль. Пожалуйста, проверьте введенные данные.")
            self.conn.commit()

    def create_change_password_form(self):
        self.clear_frame()
        frame = tbs.Frame(self.root, bootstyle="primary", padding=10)
        frame.grid(row=0, column=0, sticky=(W, E, N, S))
        tbs.Label(frame, text="Текущий пароль:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, pady=5)
        self.current_password = tk.StringVar()
        tbs.Entry(frame, textvariable=self.current_password, show="*", bootstyle="primary").grid(row=0, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Новый пароль:", bootstyle="inverse-primary").grid(row=1, column=0, sticky=W, pady=5)
        self.new_password = tk.StringVar()
        tbs.Entry(frame, textvariable=self.new_password, show="*", bootstyle="primary").grid(row=1, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Подтверждение пароля:", bootstyle="inverse-primary").grid(row=2, column=0, sticky=W, pady=5)
        self.confirm_password = tk.StringVar()
        tbs.Entry(frame, textvariable=self.confirm_password, show="*", bootstyle="primary").grid(row=2, column=1, sticky=(W, E), pady=5)
        tbs.Button(frame, text="Изменить пароль", command=self.change_password, bootstyle="primary").grid(row=3, column=0, columnspan=2, pady=10)

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
        messagebox.showinfo("Успех", "Пароль успешно изменен")
        self.create_main_menu()

    def create_main_menu(self):
        self.clear_frame()
        frame = tbs.Frame(self.root, bootstyle="primary", padding=10)
        frame.grid(row=0, column=0, sticky=(W, E, N, S))
        if self.current_user[2] == 'Администратор':
            tbs.Button(frame, text="Добавить пользователя", command=self.create_add_user_form, bootstyle="primary").grid(row=0, column=0, pady=5)
            tbs.Button(frame, text="Управление бронированиями", command=self.create_booking_form, bootstyle="primary").grid(row=1, column=0, pady=5)
            tbs.Button(frame, text="Управление номерами", command=self.create_room_management_form, bootstyle="primary").grid(row=2, column=0, pady=5)
            tbs.Button(frame, text="График уборки", command=self.create_cleaning_schedule_form, bootstyle="primary").grid(row=3, column=0, pady=5)
            tbs.Button(frame, text="Отчеты", command=self.create_reports_form, bootstyle="primary").grid(row=4, column=0, pady=5)

    def create_add_user_form(self):
        self.clear_frame()
        frame = tbs.Frame(self.root, bootstyle="primary", padding=10)
        frame.grid(row=0, column=0, sticky=(W, E, N, S))
        tbs.Label(frame, text="Имя:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, pady=5)
        self.full_name_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.full_name_var, bootstyle="primary").grid(row=0, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Логин:", bootstyle="inverse-primary").grid(row=1, column=0, sticky=W, pady=5)
        self.new_login_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.new_login_var, bootstyle="primary").grid(row=1, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Роль:", bootstyle="inverse-primary").grid(row=2, column=0, sticky=W, pady=5)
        self.role_var = tk.StringVar()
        tbs.Combobox(frame, textvariable=self.role_var, values=['Администратор', 'Руководитель', 'Уборщик'], bootstyle="primary").grid(row=2, column=1, sticky=(W, E), pady=5)
        tbs.Button(frame, text="Добавить", command=self.add_user, bootstyle="primary").grid(row=3, column=0, columnspan=2, pady=10)

    def add_user(self):
        login = self.new_login_var.get()
        self.cursor.execute("SELECT login FROM staff WHERE login = ?", (login,))
        if self.cursor.fetchone():
            messagebox.showerror("Ошибка", "Пользователь с таким логином уже существует")
            return
        self.cursor.execute("INSERT INTO staff (full_name, role, login, password) VALUES (?, ?, ?, ?)",
                          (self.full_name_var.get(), self.role_var.get(), login, self.hash_password("default")))
        self.conn.commit()
        messagebox.showinfo("Успех", "Пользователь успешно добавлен")
        self.create_main_menu()

    def create_booking_form(self):
        self.clear_frame()
        frame = tbs.Frame(self.root, bootstyle="primary", padding=10)
        frame.grid(row=0, column=0, sticky=(W, E, N, S))
        tbs.Label(frame, text="Имя гостя:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, pady=5)
        self.guest_name_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.guest_name_var, bootstyle="primary").grid(row=0, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Номер телефона:", bootstyle="inverse-primary").grid(row=1, column=0, sticky=W, pady=5)
        self.phone_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.phone_var, bootstyle="primary").grid(row=1, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Email:", bootstyle="inverse-primary").grid(row=2, column=0, sticky=W, pady=5)
        self.email_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.email_var, bootstyle="primary").grid(row=2, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Паспорт:", bootstyle="inverse-primary").grid(row=3, column=0, sticky=W, pady=5)
        self.passport_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.passport_var, bootstyle="primary").grid(row=3, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Номер комнаты:", bootstyle="inverse-primary").grid(row=4, column=0, sticky=W, pady=5)
        self.room_var = tk.StringVar()
        rooms = self.cursor.execute("SELECT room_number FROM rooms WHERE status = 'Чистый'").fetchall()
        tbs.Combobox(frame, textvariable=self.room_var, values=[r[0] for r in rooms], bootstyle="primary").grid(row=4, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Дата заезда:", bootstyle="inverse-primary").grid(row=5, column=0, sticky=W, pady=5)
        self.check_in_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.check_in_var, bootstyle="primary").grid(row=5, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Дата выезда:", bootstyle="inverse-primary").grid(row=6, column=0, sticky=W, pady=5)
        self.check_out_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.check_out_var, bootstyle="primary").grid(row=6, column=1, sticky=(W, E), pady=5)
        tbs.Button(frame, text="Забронировать", command=self.create_booking, bootstyle="primary").grid(row=7, column=0, columnspan=2, pady=10)

    def create_booking(self):
        guest_id = self.register_guest()
        room_id = self.cursor.execute("SELECT roomID FROM rooms WHERE room_number = ?", (self.room_var.get(),)).fetchone()[0]
        check_in = datetime.strptime(self.check_in_var.get(), '%Y-%m-%d').date()
        check_out = datetime.strptime(self.check_out_var.get(), '%Y-%m-%d').date()
        self.cursor.execute("INSERT INTO bookings (guest_id, room_id, check_in, check_out, status) VALUES (?, ?, ?, ?, 'Забронировано')",
                          (guest_id, room_id, check_in, check_out))
        self.cursor.execute("UPDATE rooms SET status = 'Занят' WHERE roomID = ?", (room_id,))
        self.conn.commit()
        messagebox.showinfo("Успех", "Бронирование успешно создано")
        self.create_main_menu()

    def register_guest(self):
        self.cursor.execute("INSERT OR IGNORE INTO guests (full_name, phone, email, passport) VALUES (?, ?, ?, ?)",
                          (self.guest_name_var.get(), self.phone_var.get(), self.email_var.get(), self.passport_var.get()))
        self.conn.commit()
        self.cursor.execute("SELECT guestID FROM guests WHERE passport = ?", (self.passport_var.get(),))
        return self.cursor.fetchone()[0]

    def create_room_management_form(self):
        self.clear_frame()
        frame = tbs.Frame(self.root, bootstyle="primary", padding=10)
        frame.grid(row=0, column=0, sticky=(W, E, N, S))
        tree = tbs.Treeview(frame, columns=('Номер', 'Этаж', 'Статус', 'Цена'), show='headings', bootstyle="primary")
        tree.heading('Номер', text='Номер')
        tree.heading('Этаж', text='Этаж')
        tree.heading('Статус', text='Статус')
        tree.heading('Цена', text='Цена за ночь')
        tree.grid(row=0, column=0, columnspan=2, sticky=(W, E, N, S))
        rooms = self.cursor.execute("SELECT room_number, floor, status, price_per_night FROM rooms").fetchall()
        for room in rooms:
            tree.insert('', 'end', values=room)
        tbs.Button(frame, text="Назад", command=self.create_main_menu, bootstyle="primary").grid(row=1, column=0, columnspan=2, pady=10)

    def create_cleaning_schedule_form(self):
        self.clear_frame()
        frame = tbs.Frame(self.root, bootstyle="primary", padding=10)
        frame.grid(row=0, column=0, sticky=(W, E, N, S))
        tbs.Label(frame, text="Номер:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, pady=5)
        self.cleaning_room_var = tk.StringVar()
        rooms = self.cursor.execute("SELECT room_number FROM rooms WHERE status IN ('Грязный', 'Назначен к уборке')").fetchall()
        tbs.Combobox(frame, textvariable=self.cleaning_room_var, values=[r[0] for r in rooms], bootstyle="primary").grid(row=0, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Уборщик:", bootstyle="inverse-primary").grid(row=1, column=0, sticky=W, pady=5)
        self.staff_var = tk.StringVar()
        staff = self.cursor.execute("SELECT full_name FROM staff WHERE role = 'Уборщик'").fetchall()
        tbs.Combobox(frame, textvariable=self.staff_var, values=[s[0] for s in staff], bootstyle="primary").grid(row=1, column=1, sticky=(W, E), pady=5)
        tbs.Label(frame, text="Дата уборки:", bootstyle="inverse-primary").grid(row=2, column=0, sticky=W, pady=5)
        self.cleaning_date_var = tk.StringVar()
        tbs.Entry(frame, textvariable=self.cleaning_date_var, bootstyle="primary").grid(row=2, column=1, sticky=(W, E), pady=5)
        tbs.Button(frame, text="Назначить уборку", command=self.schedule_cleaning, bootstyle="primary").grid(row=3, column=0, columnspan=2, pady=10)

    def schedule_cleaning(self):
        room_id = self.cursor.execute("SELECT roomID FROM rooms WHERE room_number = ?", (self.cleaning_room_var.get(),)).fetchone()[0]
        staff_id = self.cursor.execute("SELECT staffID FROM staff WHERE full_name = ?", (self.staff_var.get(),)).fetchone()[0]
        cleaning_date = datetime.strptime(self.cleaning_date_var.get(), '%Y-%m-%d').date()
        self.cursor.execute("INSERT INTO cleaning (room_id, staff_id, scheduled_date, status) VALUES (?, ?, ?, 'Назначено')",
                          (room_id, staff_id, cleaning_date))
        self.cursor.execute("UPDATE rooms SET status = 'Назначен к уборке' WHERE roomID = ?", (room_id,))
        self.conn.commit()
        messagebox.showinfo("Успех", "Уборка успешно назначена")
        self.create_main_menu()

    def create_reports_form(self):
        self.clear_frame()
        frame = tbs.Frame(self.root, bootstyle="primary", padding=10)
        frame.grid(row=0, column=0, sticky=(W, E, N, S))
        tbs.Label(frame, text="Дата отчета:", bootstyle="inverse-primary").grid(row=0, column=0, sticky=W, pady=5)
        self.report_date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tbs.Entry(frame, textvariable=self.report_date_var, bootstyle="primary").grid(row=0, column=1, sticky=(W, E), pady=5)
        tbs.Button(frame, text="Сформировать отчет", command=self.generate_report, bootstyle="primary").grid(row=1, column=0, columnspan=2, pady=10)

    def generate_report(self):
        report_date = datetime.strptime(self.report_date_var.get(), '%Y-%m-%d').date()
        total_rooms = self.cursor.execute("SELECT COUNT(*) FROM rooms").fetchone()[0]
        occupied_rooms = self.cursor.execute("SELECT COUNT(*) FROM bookings WHERE status IN ('Забронировано', 'Заселен') AND check_in <= ? AND check_out >= ?",
                                           (report_date, report_date)).fetchone()[0]
        occupancy_rate = (occupied_rooms / total_rooms) * 100 if total_rooms > 0 else 0
        revenue = self.cursor.execute("SELECT SUM(amount) FROM payments WHERE payment_date = ?", (report_date,)).fetchone()[0] or 0
        total_nights = self.cursor.execute("SELECT SUM(julianday(check_out) - julianday(check_in)) FROM bookings WHERE status IN ('Забронировано', 'Заселен') AND check_in <= ? AND check_out >= ?",
                                         (report_date, report_date)).fetchone()[0] or 0
        adr = revenue / total_nights if total_nights > 0 else 0
        revpar = adr * (occupancy_rate / 100)
        messagebox.showinfo("Отчет", f"Дата: {report_date}\n"
                                   f"Процент загрузки: {occupancy_rate:.2f}%\n"
                                   f"ADR: {adr:.2f}\n"
                                   f"RevPAR: {revpar:.2f}")

    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def __del__(self):
        self.conn.close()

if __name__ == "__main__":
    root = tbs.Window(themename="darkly")
    app = HotelManagementApp(root)
    root.mainloop()