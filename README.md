# Учёт актов ремонта оборудования

Бэкенд (JSON API) приложения для учёта актов ремонта офисного оборудования.
Интерфейс — отдельное React-приложение [acts-react](https://github.com/Podkolodnyi/acts-react).

## Возможности

- Инженеры: вход по паролю, регистрация с подтверждением админом
- Поиск аппаратов по серийному номеру
- Интеграция с Google Sheets (справочник аппаратов, обновляется админом)
- Интеграция с IntraService API (заявки на ремонт)
- Акты: создание, редактирование, ремонт (копия акта «Не работает»), удаление в список удалённых
- Акты ремонта узлов терморегистрации
- Автоматическая нумерация актов (инженер-месяц-порядковый номер)

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

При запуске автоматически применяются миграции базы данных (`init_db()` при загрузке `app.py`).

### 5. Интерфейс

API доступен на http://localhost:5000/api/…, интерфейс открывается через React-приложение
(адрес фронтенда задаётся в `.env` переменной `FRONTEND_ORIGIN`).

## Структура проекта

```
acts-tracking/
├── app.py                 # Основной файл приложения
├── requirements.txt       # Зависимости Python
├── start_server.bat      # Скрипт запуска для Windows
├── .env.example          # Пример переменных окружения
├── .gitignore            # Игнорируемые файлы
└── README.md             # Этот файл
```

## База данных

Приложение использует SQLite. База данных автоматически создаётся в файле `repair_acts.db` при первом запуске.

**Важно:** файл `repair_acts.db` не загружается в репозиторий (добавлен в `.gitignore`).

## Интеграции

### Google Sheets

Справочник аппаратов импортируется из Google Sheets по кнопке в админке (`POST /api/admin/devices/sync`). URL таблицы указан в `app.py`.

### IntraService

Для интеграции с IntraService необходимо указать в `.env`:
- `INTRASERVICE_BASE_URL` — адрес вашего IntraService
- `INTRASERVICE_API_LOGIN` — логин API
- `INTRASERVICE_API_PASSWORD` — пароль API

## Лицензия

Внутреннее приложение для учёта актов ремонта.
