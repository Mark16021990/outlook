```markdown
# 📧 Email Notifier - Desktop приложение для уведомлений о письмах

Приложение для автоматической проверки почты и уведомлений о новых письмах через IMAP протокол.

## 📋 Функциональность

- ✅ Автоматическая проверка почты по расписанию
- ✅ Уведомления о новых письмах
- ✅ Отображение списка последних писем
- ✅ Безопасное подключение через SSL
- ✅ Современный графический интерфейс
- ✅ Настройка интервала проверки

## 🚀 Быстрый старт

### Предварительные требования

- Python 3.8+
- Доступ к почтовому серверу по IMAP
- Учетные данные для подключения к почте

### Установка и запуск

1. **Клонирование репозитория**
   ```bash
   git clone https://github.com/Mark16021990/outlook.git
   cd outlook
   ```

2. **Создание виртуального окружения**
   ```bash
   python -m venv venv
   ```

3. **Активация виртуального окружения**
   - Windows:
     ```bash
     venv\Scripts\activate
     ```
   - Linux/MacOS:
     ```bash
     source venv/bin/activate
     ```

4. **Установка зависимостей**
   ```bash
   pip install -r requirements.txt
   ```

5. **Запуск приложения**
   ```bash
   python main.py
   ```

### Конфигурация

По умолчанию приложение использует следующие учетные данные:
- Хост: `mail.evgrp.ru`
- Логин: `tk/ma.andrejchenko`
- Пароль: `Road90-Fill5`
- Почта: `ma.andrejchenko@stroyenergokom.ru`

Для изменения настроек отредактируйте соответствующие поля в файле `main.py`.

## 📦 Сборка в исполняемый файл

### Создание EXE-файла для Windows

1. **Убедитесь, что установлены все зависимости**
   ```bash
   pip install pyinstaller
   ```

2. **Сборка приложения**
   ```bash
   python build.py
   ```

3. **Готовый файл будет расположен в папке `dist/EmailNotifier.exe`**

### Ручная сборка с PyInstaller

```bash
pyinstaller --onefile --windowed --name=EmailNotifier --hidden-import=imapclient --hidden-import=PIL --hidden-import=customtkinter --hidden-import=email --hidden-import=ssl --hidden-import=CTkMessagebox main.py
```

## ⚙️ Настройки приложения

### Изменение интервала проверки

1. Запустите приложение
2. Нажмите кнопку "Настройки"
3. Укажите желаемый интервал проверки в секундах
4. Сохраните настройки

### Изменение почтовых учетных данных

Для изменения учетных данных отредактируйте следующие переменные в файле `main.py`:

```python
self.host = "mail.evgrp.ru"
self.username = "tk/ma.andrejchenko"
self.password = "Road90-Fill5"
self.email_address = "ma.andrejchenko@stroyenergokom.ru"
```

## 🔧 Устранение неполадок

### Проблемы с подключением

1. **Проверьте интернет-соединение**
2. **Убедитесь в правильности учетных данных**
3. **Проверьте настройки firewall/антивируса**

### Проблемы с SSL сертификатами

Если возникают ошибки SSL, попробуйте отключить проверку сертификата:

В файле `main.py` в методе `connect_to_server` замените:
```python
ssl_context = ssl.create_default_context()
```
на:
```python
ssl_context = ssl._create_unverified_context()
```

### Проблемы со сборкой

Если сборка с PyInstaller завершается с ошибками:
1. Убедитесь, что все зависимости установлены
2. Попробуйте собрать с дополнительными скрытыми импортами:
   ```bash
   pyinstaller --onefile --windowed --name=EmailNotifier --hidden-import=imapclient --hidden-import=PIL --hidden-import=customtkinter --hidden-import=email --hidden-import=ssl --hidden-import=CTkMessagebox --hidden-import=email.header --hidden-import=email.policy main.py
   ```

## 📁 Структура проекта

```
outlook/
├── main.py              # Основной файл приложения
├── build.py             # Скрипт для сборки
├── requirements.txt     # Зависимости проекта
├── README.md           # Документация
├── assets/             # Ресурсы приложения
└── dist/               # Собранные исполняемые файлы
```

## 🛠️ Технологии

- **Python** - основной язык программирования
- **CustomTkinter** - современная графическая библиотека
- **IMAPClient** - работа с IMAP протоколом
- **PyInstaller** - сборка в исполняемый файл

## 📄 Лицензия

Проект распространяется под лицензией MIT. Подробнее см. в файле LICENSE.

## 🤝 Поддержка

Если у вас возникли вопросы или проблемы:
1. Проверьте документацию и раздел "Устранение неполадок"
2. Создайте issue в репозитории проекта
3. Опишите подробно вашу проблему и шаги для ее воспроизведения

## 🔄 Обновления

Для обновления приложения до последней версии:
```bash
git pull origin main
pip install -r requirements.txt --upgrade
```

---

**Примечание**: Убедитесь, что использование приложения соответствует политике безопасности вашей организации и правилам использования почтового сервера.
```