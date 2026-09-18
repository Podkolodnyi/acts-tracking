# Учёт актов ремонта оборудования

Веб-приложение для учёта актов ремонта офисного оборудования.

## Возможности

- Учёт инженеров (Подколодный, Дзюба, Изъюров, Озеров)
- Поиск аппаратов по серийному номеру
- Интеграция с Google Sheets (справочник аппаратов)
- Интеграция с IntraService API (заявки на ремонт)
- Создание и редактирование актов
- Черновики и завершённые акты
- Автоматическая нумерация актов (инженер-месяц-порядковый номер)
- Выгрузка актов в Excel
- Печатная форма акта

## Быстрый старт

### 1. Клонирование репозитория

```bash
git clone https://github.com/Podkolodnyi/acts-tracking.git
cd acts-tracking
```

### 2. Установка зависимостей

```bash
pip install -r requirements.txt
```

### 3. Настройка переменных окружения

Создайте файл `.env` на основе `.env.example`:

```bash
copy .env.example .env
```

Заполните `.env`:

```env
# IntraService API
INTRASERVICE_BASE_URL=https://your-intraservice-instance.ru
INTRASERVICE_API_LOGIN=your_api_login
INTRASERVICE_API_PASSWORD=your_api_password

# Flask
FLASK_SECRET_KEY=change-this-before-network-use
```

### 4. Запуск сервера

**Вариант A: через start_server.bat (Windows)**

```bash
start_server.bat
```

**Вариант B: вручную**

```bash
python -m waitress --listen=0.0.0.0:5000 app:app
```

**Вариант C: режим разработки**

```bash
python app.py
```

### 5. Открыть в браузере

Перейдите по адресу: http://localhost:5000

## Структура проекта

```
acts-tracking/
├── app.py                 # Основной файл приложения
├── requirements.txt       # Зависимости Python
├── start_server.bat      # Скрипт запуска для Windows
├── .env.example          # Пример переменных окружения
├── .gitignore            # Игнорируемые файлы
├── README.md             # Этот файл
└── templates/            # HTML-шаблоны
    ├── base.html         # Базовый шаблон
    ├── engineer.html     # Выбор инженера
    ├── home.html         # Главная страница
    ├── devices.html      # Поиск аппаратов
    ├── acts.html         # Список актов
    ├── act_form.html     # Форма создания/редактирования акта
    └── act_print.html    # Печатная форма акта
```

## База данных

Приложение использует SQLite. База данных автоматически создаётся в файле `repair_acts.db` при первом запуске.

**Важно:** файл `repair_acts.db` не загружается в репозиторий (добавлен в `.gitignore`).

## Интеграции

### Google Sheets

Приложение автоматически импортирует справочник аппаратов из Google Sheets при запуске. URL таблицы указан в `app.py`.

### IntraService

Для интеграции с IntraService необходимо указать в `.env`:
- `INTRASERVICE_BASE_URL` — адрес вашего IntraService
- `INTRASERVICE_API_LOGIN` — логин API
- `INTRASERVICE_API_PASSWORD` — пароль API

## Добавление API для React

Для подключения React-фронтенда необходимо добавить REST API endpoint'ы в `app.py`:

- `GET /api/engineers` — список инженеров
- `GET /api/devices?q=...` — поиск аппаратов
- `GET /api/acts` — список актов с фильтрами
- `POST /api/acts` — создание акта
- `GET /api/acts/<id>` — получение акта
- `PUT /api/acts/<id>` — редактирование акта
- `DELETE /api/acts/<id>` — удаление акта

## Лицензия

Внутреннее приложение для учёта актов ремонта.
