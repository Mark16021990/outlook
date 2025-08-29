import customtkinter as ctk
import threading
import time
from imapclient import IMAPClient
import email
from email.header import decode_header
import ssl
from CTkMessagebox import CTkMessagebox
import os
from PIL import Image
import sys

# Настройка внешнего вида
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class EmailNotifier(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Настройка окна
        self.title("📧 Email Notifier")
        self.geometry("600x700")
        self.minsize(500, 600)
        
        # Данные для подключения
        self.host = "mail.evgrp.ru"
        self.username = "tk/ma.andrejchenko"
        self.password = "Road90-Fill5"
        self.email_address = "ma.andrejchenko@stroyenergokom.ru"
        
        # Переменные состояния
        self.connected = False
        self.client = None
        self.last_check = None
        self.unread_count = 0
        self.check_interval = 300  # 5 минут
        
        # ID последних проверенных писем
        self.last_checked_ids = set()
        
        # Создание интерфейса
        self.setup_ui()
        
        # Запуск проверки почты
        self.check_emails()
        
    def setup_ui(self):
        # Основной фрейм
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Заголовок
        self.title_label = ctk.CTkLabel(self.main_frame, text="📧 Email Notifier", 
                                      font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(pady=20)
        
        # Статус подключения
        self.status_frame = ctk.CTkFrame(self.main_frame)
        self.status_frame.pack(fill="x", padx=20, pady=10)
        
        self.status_label = ctk.CTkLabel(self.status_frame, text="Статус: Отключено", 
                                       font=ctk.CTkFont(size=14))
        self.status_label.pack(side="left", padx=10)
        
        self.connection_indicator = ctk.CTkLabel(self.status_frame, text="●", 
                                               text_color="red", font=ctk.CTkFont(size=16))
        self.connection_indicator.pack(side="right", padx=10)
        
        # Информация о почте
        self.info_frame = ctk.CTkFrame(self.main_frame)
        self.info_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(self.info_frame, text="Почта:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.email_label = ctk.CTkLabel(self.info_frame, text=self.email_address, 
                                      font=ctk.CTkFont(weight="bold"))
        self.email_label.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        ctk.CTkLabel(self.info_frame, text="Непрочитанные:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.unread_label = ctk.CTkLabel(self.info_frame, text="0", 
                                       font=ctk.CTkFont(weight="bold"))
        self.unread_label.grid(row=1, column=1, sticky="w", padx=10, pady=5)
        
        ctk.CTkLabel(self.info_frame, text="Последняя проверка:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.last_check_label = ctk.CTkLabel(self.info_frame, text="Никогда", 
                                           font=ctk.CTkFont(weight="bold"))
        self.last_check_label.grid(row=2, column=1, sticky="w", padx=10, pady=5)
        
        # Кнопки управления
        self.buttons_frame = ctk.CTkFrame(self.main_frame)
        self.buttons_frame.pack(fill="x", padx=20, pady=10)
        
        self.check_button = ctk.CTkButton(self.buttons_frame, text="Проверить сейчас", 
                                        command=self.check_emails)
        self.check_button.pack(side="left", padx=5)
        
        self.settings_button = ctk.CTkButton(self.buttons_frame, text="Настройки", 
                                           command=self.open_settings)
        self.settings_button.pack(side="right", padx=5)
        
        # Список писем
        self.emails_frame = ctk.CTkFrame(self.main_frame)
        self.emails_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        ctk.CTkLabel(self.emails_frame, text="Последние письма:", 
                   font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        # Таблица с письмами
        self.emails_table = ctk.CTkTextbox(self.emails_frame, height=300)
        self.emails_table.pack(fill="both", expand=True, padx=10, pady=10)
        self.emails_table.configure(state="disabled")
        
        # Статус бар
        self.status_bar = ctk.CTkLabel(self, text="Готов к работе", anchor="w")
        self.status_bar.pack(fill="x", side="bottom", padx=20, pady=5)
    
    def connect_to_server(self):
        """Подключение к IMAP серверу"""
        try:
            # Создаем SSL контекст
            ssl_context = ssl.create_default_context()
            
            # Подключаемся к серверу
            self.client = IMAPClient(self.host, ssl_context=ssl_context)
            self.client.login(self.username, self.password)
            
            # Выбираем папку входящих
            self.client.select_folder('INBOX')
            
            self.connected = True
            self.update_status("Подключено", "green")
            return True
            
        except Exception as e:
            self.connected = False
            self.update_status(f"Ошибка подключения: {str(e)}", "red")
            return False
    
    def disconnect_from_server(self):
        """Отключение от сервера"""
        try:
            if self.client:
                self.client.logout()
                self.client = None
        except:
            pass
        finally:
            self.connected = False
            self.update_status("Отключено", "red")
    
    def check_emails(self):
        """Проверка новых писем"""
        # Обновляем статус
        self.update_status("Проверка почты...", "orange")
        self.check_button.configure(state="disabled")
        
        # Запускаем в отдельном потоке
        thread = threading.Thread(target=self._check_emails_thread)
        thread.daemon = True
        thread.start()
    
    def _check_emails_thread(self):
        """Поток для проверки почты"""
        try:
            # Подключаемся к серверу
            if not self.connect_to_server():
                return
            
            # Ищем непрочитанные письма
            messages = self.client.search(['UNSEEN'])
            self.unread_count = len(messages)
            
            # Получаем информацию о письмах
            emails_info = []
            new_emails = []
            
            for msg_id in messages:
                # Проверяем,是新ое ли это письмо
                if msg_id not in self.last_checked_ids:
                    new_emails.append(msg_id)
                
                # Получаем данные письма
                msg_data = self.client.fetch([msg_id], ['RFC822', 'BODY.PEEK[]'])
                raw_email = msg_data[msg_id][b'RFC822']
                
                # Парсим письмо
                email_message = email.message_from_bytes(raw_email)
                
                # Получаем тему
                subject, encoding = decode_header(email_message["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")
                
                # Получаем отправителя
                from_ = email_message.get("From")
                
                # Получаем дату
                date = email_message.get("Date")
                
                emails_info.append({
                    'id': msg_id,
                    'subject': subject,
                    'from': from_,
                    'date': date
                })
            
            # Обновляем последние проверенные ID
            self.last_checked_ids = set(messages)
            
            # Обновляем интерфейс в главном потоке
            self.after(0, lambda: self.update_emails_list(emails_info, new_emails))
            
            # Показываем уведомления о новых письмах
            if new_emails:
                self.after(0, lambda: self.show_new_emails_notification(len(new_emails)))
                
        except Exception as e:
            self.after(0, lambda: self.update_status(f"Ошибка: {str(e)}", "red"))
        finally:
            # Отключаемся от сервера
            self.disconnect_from_server()
            
            # Запускаем следующую проверку через интервал
            self.after(self.check_interval * 1000, self.check_emails)
            
            # Обновляем статус
            self.after(0, lambda: self.update_status("Готов к работе", "white"))
            self.after(0, lambda: self.check_button.configure(state="normal"))
    
    def update_emails_list(self, emails_info, new_emails):
        """Обновление списка писем"""
        # Обновляем счетчик непрочитанных
        self.unread_label.configure(text=str(self.unread_count))
        
        # Обновляем время последней проверки
        self.last_check = time.strftime("%H:%M:%S")
        self.last_check_label.configure(text=self.last_check)
        
        # Обновляем список писем
        self.emails_table.configure(state="normal")
        self.emails_table.delete("1.0", "end")
        
        for email_info in emails_info[-10:]:  # Показываем последние 10 писем
            is_new = email_info['id'] in new_emails
            new_indicator = "🆕 " if is_new else ""
            
            email_text = f"{new_indicator}От: {email_info['from']}\n"
            email_text += f"Тема: {email_info['subject']}\n"
            email_text += f"Дата: {email_info['date']}\n"
            email_text += "-" * 50 + "\n"
            
            self.emails_table.insert("end", email_text)
        
        self.emails_table.configure(state="disabled")
    
    def show_new_emails_notification(self, count):
        """Показ уведомления о новых письмах"""
        CTkMessagebox(title="Новые письма", 
                     message=f"У вас {count} новых письма(ем)!", 
                     icon="info")
    
    def update_status(self, message, color="white"):
        """Обновление статуса"""
        self.status_label.configure(text=f"Статус: {message}")
        self.connection_indicator.configure(text_color=color)
        self.status_bar.configure(text=message)
    
    def open_settings(self):
        """Открытие окна настроек"""
        settings_window = ctk.CTkToplevel(self)
        settings_window.title("Настройки")
        settings_window.geometry("400x300")
        settings_window.transient(self)
        settings_window.grab_set()
        
        # Настройки интервала проверки
        ctk.CTkLabel(settings_window, text="Интервал проверки (секунды):").pack(pady=10)
        interval_var = ctk.StringVar(value=str(self.check_interval))
        interval_entry = ctk.CTkEntry(settings_window, textvariable=interval_var)
        interval_entry.pack(pady=5)
        
        # Кнопка сохранения
        def save_settings():
            try:
                new_interval = int(interval_var.get())
                if new_interval < 30:
                    CTkMessagebox(title="Ошибка", message="Интервал не может быть меньше 30 секунд", icon="warning")
                    return
                
                self.check_interval = new_interval
                settings_window.destroy()
                CTkMessagebox(title="Успех", message="Настройки сохранены", icon="info")
            except ValueError:
                CTkMessagebox(title="Ошибка", message="Введите корректное число", icon="warning")
        
        ctk.CTkButton(settings_window, text="Сохранить", command=save_settings).pack(pady=20)
    
    def on_closing(self):
        """Обработчик закрытия приложения"""
        self.disconnect_from_server()
        self.destroy()

if __name__ == "__main__":
    app = EmailNotifier()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()