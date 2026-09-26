
from flask_cors import CORS
import csv
import io
import json
import os
import sqlite3
import urllib.request
from contextlib import contextmanager
from datetime import datetime, timedelta

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, request, session
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

INTRASERVICE_BASE_URL = os.environ["INTRASERVICE_BASE_URL"].rstrip("/")
INTRASERVICE_AUTH = (
    os.environ["INTRASERVICE_API_LOGIN"],
    os.environ["INTRASERVICE_API_PASSWORD"],
)

DATABASE = os.path.join(BASE_DIR, "repair_acts.db")
SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1wq-TfC-wGvyZTFTQLSbsxles7KJ3Owj3KYXeiAcoeP8/"
    "export?format=csv&gid=467885286"
)

# One-time seed data for the `engineers` table (see init_db). Editing this
# after the first run has no effect — it only backfills an empty table.
SEED_ENGINEERS = {
    "podkolodny": {"name": "Подколодный", "code": "1", "is_admin": True},
    "dzyuba": {"name": "Дзюба", "code": "2", "is_admin": False},
    "izyurov": {"name": "Изъюров", "code": "5", "is_admin": False},
    "ozerov": {"name": "Озеров", "code": "9", "is_admin": False},
}

# This account can never be removed by another admin.
PROTECTED_ENGINEER_KEY = "podkolodny"

NUMBER_TYPES = {
    "standard": {"name": "Обычный акт"},
    "thermo": {"name": "Узел терморегистрации"},
}

# Act state is the device condition. It is set once when the act is created
# and never changes; a broken device is "repaired" by creating a copy act.
CONDITION_WORKING = "Работает"
CONDITION_BROKEN = "Не работает"
DEVICE_CONDITIONS = {CONDITION_WORKING, CONDITION_BROKEN}

MATERIAL_PREFIXES = [
    ("Наименование:", "name"),
    ("Парт:", "article"),
    ("Артикул:", "article"),
    ("Кол-во:", "quantity"),
]

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get(
    "FLASK_SECRET_KEY",
    "change-this-before-network-use",
)
# Enable CORS for React frontend (with credentials, so the session cookie works)
FRONTEND_ORIGIN = os.environ.get("FRONTEND_ORIGIN", "http://localhost:5173")
CORS(app, supports_credentials=True, origins=[FRONTEND_ORIGIN])


@contextmanager
def db():
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    try:
        yield connection
        connection.commit()
    finally:
        connection.close()


def clean(value):
    return " ".join((value or "").strip().split())


def clean_multiline(value):
    lines = []
    for line in (value or "").splitlines():
        line = " ".join(line.strip().split())
        if line:
            lines.append(line)
    return "\n".join(lines)


def normalize_serial(value):
    return clean(value).upper().replace(" ", "")


def parse_materials_text(text):
    """Parse legacy "Наименование: X; Артикул: Y; Кол-во: Z" lines."""
    items = []
    for line in (text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        item = {"name": "", "article": "", "quantity": ""}
        for chunk in line.split(";"):
            part = chunk.strip()
            for prefix, key in MATERIAL_PREFIXES:
                if part.startswith(prefix):
                    item[key] = part[len(prefix):].strip()
                    break
        items.append(item)
    return items


def save_act_materials(con, act_id, materials):
    """Replace all materials of an act. Rows without any value are skipped."""
    con.execute("DELETE FROM act_materials WHERE act_id = ?", (act_id,))
    position = 0
    for item in materials:
        name = clean(item.get("name"))
        article = clean(item.get("article"))
        quantity = clean(item.get("quantity"))
        if not (name or article or quantity):
            continue
        position += 1
        con.execute(
            "INSERT INTO act_materials "
            "(act_id, position, name, article, quantity) "
            "VALUES (?, ?, ?, ?, ?)",
            (act_id, position, name, article, quantity),
        )


def get_act_materials(con, act_id):
    rows = con.execute(
        "SELECT name, article, quantity FROM act_materials "
        "WHERE act_id = ? ORDER BY position",
        (act_id,),
    ).fetchall()
    return [dict(row) for row in rows]


def field(row, *names):
    lower = {clean(key).lower(): value for key, value in row.items()}
    for name in names:
        value = lower.get(name.lower())
        if value is not None:
            return clean(value)
    return ""


def search_intraservice(query):
    query = clean(query)
    if not query:
        return []

    response = requests.get(
        f"{INTRASERVICE_BASE_URL}/api/task",
        params={"pagesize": 20, "archive": "true", "search": query},
        auth=INTRASERVICE_AUTH,
        headers={"Accept": "application/json"},
        timeout=20,
    )
    response.raise_for_status()
    payload = response.json()

    return [
        {
            "id": task.get("Id"),
            "name": task.get("Name") or "",
            "description": task.get("Description") or "",
            "created": task.get("Created") or "",
            "deadline": task.get("Deadline") or "",
        }
        for task in payload.get("Tasks", [])
    ]


def init_db():
    with db() as con:
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS devices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source TEXT NOT NULL,
                source_key TEXT NOT NULL,
                sap_id TEXT,
                mvz TEXT,
                customer_name TEXT NOT NULL,
                object_format TEXT,
                address TEXT,
                model TEXT,
                serial_number TEXT NOT NULL,
                serial_normalized TEXT NOT NULL,
                imported_at TEXT NOT NULL,
                UNIQUE(source, source_key)
            );

            CREATE INDEX IF NOT EXISTS idx_devices_serial
                ON devices(serial_normalized);
            CREATE INDEX IF NOT EXISTS idx_devices_customer
                ON devices(customer_name);

            CREATE TABLE IF NOT EXISTS acts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                act_number TEXT UNIQUE,
                number_type TEXT NOT NULL DEFAULT 'standard',
                engineer_key TEXT NOT NULL,
                engineer_name TEXT NOT NULL,
                engineer_code TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'draft',
                customer_name TEXT,
                customer_representative TEXT,
                device_model TEXT,
                serial_number TEXT,
                printeco_label TEXT,
                device_type TEXT,
                comment_label TEXT,
                address TEXT,
                phone TEXT,
                fault TEXT,
                counter_bw TEXT,
                counter_color TEXT,
                service_kind TEXT,
                diagnostics_result TEXT,
                works_text TEXT,
                materials_text TEXT,
                work_date TEXT,
                start_time TEXT,
                end_time TEXT,
                device_condition TEXT,
                customer_signatory TEXT,
                source_device_id INTEGER,
                intraservice_task_id TEXT,
                FOREIGN KEY(source_device_id) REFERENCES devices(id)
            );

            CREATE INDEX IF NOT EXISTS idx_acts_serial
                ON acts(serial_number);
            CREATE INDEX IF NOT EXISTS idx_acts_number
                ON acts(act_number);
            CREATE INDEX IF NOT EXISTS idx_acts_engineer
                ON acts(engineer_key);

            CREATE TABLE IF NOT EXISTS engineers (
                key TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                code TEXT NOT NULL UNIQUE,
                password_hash TEXT,
                is_admin INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL DEFAULT 'active',
                created_at TEXT NOT NULL
            );
            """
        )

        columns = {
            row["name"]
            for row in con.execute("PRAGMA table_info(acts)").fetchall()
        }
        if "number_type" not in columns:
            con.execute(
                "ALTER TABLE acts "
                "ADD COLUMN number_type TEXT NOT NULL DEFAULT 'standard'"
            )
        # Cross-links between a broken act and its repaired copy.
        if "repaired_by_act_id" not in columns:
            con.execute(
                "ALTER TABLE acts ADD COLUMN repaired_by_act_id INTEGER "
                "REFERENCES acts(id)"
            )
        if "repair_of_act_id" not in columns:
            con.execute(
                "ALTER TABLE acts ADD COLUMN repair_of_act_id INTEGER "
                "REFERENCES acts(id)"
            )

        # Soft delete: a deleted act gets a DEL-NNNN number, its old number
        # is kept in original_act_number and becomes free for new acts.
        for column in ("deleted_at", "deleted_by", "original_act_number"):
            if column not in columns:
                con.execute(f"ALTER TABLE acts ADD COLUMN {column} TEXT")

        # Counters that must never go back, even after rows are purged.
        con.execute(
            "CREATE TABLE IF NOT EXISTS sequences ("
            "name TEXT PRIMARY KEY, value INTEGER NOT NULL)"
        )

        # Previous states of edited acts (admin-only "Версии актов").
        # data is a JSON snapshot in the same shape as GET /api/acts/<id>.
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS act_versions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                act_id INTEGER NOT NULL REFERENCES acts(id),
                version_number TEXT NOT NULL,
                created_at TEXT NOT NULL,
                created_by TEXT NOT NULL,
                data TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_act_versions_act
                ON act_versions(act_id);
            """
        )

        # Legacy acts were saved without a condition; they are all working.
        # Thermo acts have no condition at all, so they are left alone.
        con.execute(
            "UPDATE acts SET device_condition = ? "
            "WHERE (device_condition IS NULL OR device_condition = '') "
            "AND number_type = 'standard'",
            (CONDITION_WORKING,),
        )
        # work_date is required now; legacy acts without it get the date
        # the act was created.
        con.execute(
            "UPDATE acts SET work_date = DATE(created_at) "
            "WHERE work_date IS NULL OR work_date = ''"
        )

        has_materials_table = con.execute(
            "SELECT 1 FROM sqlite_master "
            "WHERE type = 'table' AND name = 'act_materials'"
        ).fetchone()
        con.executescript(
            """
            CREATE TABLE IF NOT EXISTS act_materials (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                act_id INTEGER NOT NULL REFERENCES acts(id),
                position INTEGER NOT NULL,
                name TEXT NOT NULL DEFAULT '',
                article TEXT NOT NULL DEFAULT '',
                quantity TEXT NOT NULL DEFAULT ''
            );

            CREATE INDEX IF NOT EXISTS idx_act_materials_act
                ON act_materials(act_id);
            """
        )
        # One-time backfill from the old multi-line materials_text column.
        if not has_materials_table:
            rows = con.execute(
                "SELECT id, materials_text FROM acts "
                "WHERE materials_text IS NOT NULL AND materials_text <> ''"
            ).fetchall()
            for row in rows:
                save_act_materials(
                    con, row["id"], parse_materials_text(row["materials_text"])
                )

        engineer_columns = {
            row["name"]
            for row in con.execute("PRAGMA table_info(engineers)").fetchall()
        }
        if "first_name" not in engineer_columns:
            con.execute("ALTER TABLE engineers ADD COLUMN first_name TEXT")

        has_engineers = con.execute(
            "SELECT 1 FROM engineers LIMIT 1"
        ).fetchone()
        if not has_engineers:
            now = datetime.now().isoformat(timespec="seconds")
            con.executemany(
                "INSERT INTO engineers "
                "(key, name, code, password_hash, is_admin, status, created_at) "
                "VALUES (?, ?, ?, NULL, ?, 'active', ?)",
                [
                    (key, value["name"], value["code"], int(value["is_admin"]), now)
                    for key, value in SEED_ENGINEERS.items()
                ],
            )


def get_engineers_dict(con=None):
    if con is None:
        with db() as fresh_con:
            return get_engineers_dict(fresh_con)

    rows = con.execute(
        "SELECT key, name, code, is_admin, status FROM engineers"
    ).fetchall()
    return {
        row["key"]: {
            "name": row["name"],
            "code": row["code"],
            "is_admin": bool(row["is_admin"]),
            "status": row["status"],
        }
        for row in rows
    }


def get_engineer():
    key = session.get("engineer_key")
    if not key:
        return None

    with db() as con:
        row = con.execute(
            "SELECT name, first_name, code, is_admin FROM engineers "
            "WHERE key = ? AND status = 'active'",
            (key,),
        ).fetchone()

    if not row:
        return None

    return {
        "name": row["name"],
        "first_name": row["first_name"] or "",
        "code": row["code"],
        "is_admin": bool(row["is_admin"]),
    }


def require_admin():
    engineer = get_engineer()
    return bool(engineer and engineer["is_admin"])


CYRILLIC_TO_LATIN = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
    "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "kh", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "shch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
}


def transliterate(text):
    return "".join(
        CYRILLIC_TO_LATIN.get(ch, ch if ch.isalnum() else "")
        for ch in text.lower()
    )


def generate_engineer_key(con, first_name, last_name):
    def is_free(candidate):
        return not con.execute(
            "SELECT 1 FROM engineers WHERE key = ?", (candidate,)
        ).fetchone()

    base = transliterate(last_name)
    if base and is_free(base):
        return base

    combined = f"{base}_{transliterate(first_name)}".strip("_")
    if combined and is_free(combined):
        return combined

    fallback = combined or base or "engineer"
    suffix = 2
    while not is_free(f"{fallback}{suffix}"):
        suffix += 1
    return f"{fallback}{suffix}"


def generate_engineer_code(con):
    used_codes = {
        int(row["code"])
        for row in con.execute("SELECT code FROM engineers").fetchall()
        if row["code"].isdigit()
    }
    candidate = 10
    while candidate in used_codes:
        candidate += 1
    return str(candidate)


def next_act_number(
    con,
    engineer_code,
    now,
    condition=CONDITION_WORKING,
    number_type="standard",
):
    month = now.strftime("%m")

    # Thermo acts have no condition, so they are always "T-".
    if number_type == "thermo":
        prefix = f"T-{engineer_code}-{month}-"
    elif condition == CONDITION_BROKEN:
        prefix = f"D-{engineer_code}-{month}-"
    else:
        prefix = f"{engineer_code}-{month}-"

    row = con.execute(
        """
        SELECT act_number
        FROM acts
        WHERE act_number LIKE ?
        ORDER BY act_number DESC
        LIMIT 1
        """,
        (prefix + "%",),
    ).fetchone()

    sequence = (
        int(row["act_number"].rsplit("-", 1)[-1]) + 1
        if row
        else 1
    )

    return f"{prefix}{sequence:04d}"


def import_google_sheet():
    request_obj = urllib.request.Request(
        SHEET_URL,
        headers={"User-Agent": "RepairActs/1.0"},
    )

    with urllib.request.urlopen(request_obj, timeout=30) as response:
        content = response.read().decode("utf-8-sig")

    reader = csv.DictReader(io.StringIO(content))
    if not reader.fieldnames:
        raise ValueError("В Google Sheets не найдена строка заголовков.")

    now = datetime.now().isoformat(timespec="seconds")
    imported = 0

    with db() as con:
        for line_number, row in enumerate(reader, start=2):
            sap_id = field(row, "SAP")
            mvz = field(row, "МВЗ", "MB3", "МВЗ ")
            customer = field(row, "клиент", "Клиент")
            object_format = field(row, "Формат объект", "Формат объекта")
            address = field(row, "Адрес объекта")
            model = field(row, "Модель принтера", "Модель принт", "Модель")
            serial = field(row, "Серийный номер", "Серийный ном", "Серийник")

            if not serial or not customer:
                continue

            con.execute(
                """
                INSERT INTO devices (
                    source, source_key, sap_id, mvz, customer_name,
                    object_format, address, model, serial_number,
                    serial_normalized, imported_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(source, source_key) DO UPDATE SET
                    sap_id=excluded.sap_id,
                    mvz=excluded.mvz,
                    customer_name=excluded.customer_name,
                    object_format=excluded.object_format,
                    address=excluded.address,
                    model=excluded.model,
                    serial_number=excluded.serial_number,
                    serial_normalized=excluded.serial_normalized,
                    imported_at=excluded.imported_at
                """,
                (
                    "google_sheet_main",
                    str(line_number),
                    sap_id,
                    mvz,
                    customer,
                    object_format,
                    address,
                    model,
                    serial,
                    normalize_serial(serial),
                    now,
                ),
            )
            imported += 1

    return imported


def save_manual_device(con, values, now):
    serial = clean(values.get("serial_number"))
    customer = clean(values.get("customer_name"))

    if not serial or not customer:
        return None

    serial_normalized = normalize_serial(serial)
    existing = con.execute(
        """
        SELECT id
        FROM devices
        WHERE source = 'manual' AND serial_normalized = ?
        LIMIT 1
        """,
        (serial_normalized,),
    ).fetchone()

    data = (
        customer,
        clean(values.get("address")),
        clean(values.get("device_model")),
        serial,
        serial_normalized,
        now.isoformat(timespec="seconds"),
    )

    if existing:
        con.execute(
            """
            UPDATE devices SET
                customer_name=?,
                address=?,
                model=?,
                serial_number=?,
                serial_normalized=?,
                imported_at=?
            WHERE id=?
            """,
            data + (existing["id"],),
        )
        return existing["id"]

    cursor = con.execute(
        """
        INSERT INTO devices (
            source, source_key, sap_id, mvz, customer_name,
            object_format, address, model, serial_number,
            serial_normalized, imported_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            "manual",
            f"manual:{serial_normalized}",
            "",
            "",
            customer,
            "",
            clean(values.get("address")),
            clean(values.get("device_model")),
            serial,
            serial_normalized,
            now.isoformat(timespec="seconds"),
        ),
    )
    return cursor.lastrowid


@app.route("/api/session", methods=["GET"])
def api_session():
    engineer = get_engineer()
    if not engineer:
        return jsonify({"engineer": None})

    return jsonify({
        "engineer": {
            "id": session.get("engineer_key"),
            "name": engineer["name"],
            "first_name": engineer["first_name"],
            "code": engineer["code"],
            "is_admin": engineer["is_admin"],
        }
    })


@app.route("/api/session", methods=["DELETE"])
def api_session_logout():
    session.clear()
    return jsonify({"engineer": None})


@app.route("/api/session/claim", methods=["POST"])
def api_session_claim():
    data = request.get_json(silent=True) or {}
    key = data.get("engineer_key", "")
    first_name = clean(data.get("first_name", ""))
    password = data.get("password", "")

    if not first_name:
        return jsonify({
            "error": "Укажите имя",
            "code": "MISSING_FIELDS",
        }), 400

    if len(password) < 6:
        return jsonify({
            "error": "Пароль должен быть не короче 6 символов",
            "code": "WEAK_PASSWORD",
        }), 400

    with db() as con:
        row = con.execute(
            "SELECT key, name, code, password_hash, is_admin FROM engineers "
            "WHERE key = ? AND status = 'active'",
            (key,),
        ).fetchone()

        if not row:
            return jsonify({
                "error": "Инженер не найден",
                "code": "NOT_FOUND",
            }), 404

        if row["password_hash"] is not None:
            return jsonify({
                "error": "Пароль уже задан, используйте вход",
                "code": "ALREADY_CLAIMED",
            }), 400

        con.execute(
            "UPDATE engineers SET password_hash = ?, first_name = ? "
            "WHERE key = ?",
            (generate_password_hash(password), first_name, key),
        )

    session["engineer_key"] = key
    return jsonify({
        "engineer": {
            "id": row["key"],
            "name": row["name"],
            "first_name": first_name,
            "code": row["code"],
            "is_admin": bool(row["is_admin"]),
        }
    })


@app.route("/api/session/login", methods=["POST"])
def api_session_login():
    data = request.get_json(silent=True) or {}
    key = data.get("engineer_key", "")
    password = data.get("password", "")

    with db() as con:
        row = con.execute(
            "SELECT key, name, first_name, code, password_hash, is_admin "
            "FROM engineers WHERE key = ? AND status = 'active'",
            (key,),
        ).fetchone()

    if not row or not row["password_hash"]:
        return jsonify({
            "error": "Неверный логин или пароль",
            "code": "INVALID_CREDENTIALS",
        }), 401

    if not check_password_hash(row["password_hash"], password):
        return jsonify({
            "error": "Неверный логин или пароль",
            "code": "INVALID_CREDENTIALS",
        }), 401

    session["engineer_key"] = key
    return jsonify({
        "engineer": {
            "id": row["key"],
            "name": row["name"],
            "first_name": row["first_name"] or "",
            "code": row["code"],
            "is_admin": bool(row["is_admin"]),
        }
    })


@app.route("/api/session/register", methods=["POST"])
def api_session_register():
    data = request.get_json(silent=True) or {}
    first_name = clean(data.get("first_name", ""))
    last_name = clean(data.get("last_name", ""))
    password = data.get("password", "")

    if not first_name or not last_name:
        return jsonify({
            "error": "Заполните все поля",
            "code": "MISSING_FIELDS",
        }), 400

    if len(password) < 6:
        return jsonify({
            "error": "Пароль должен быть не короче 6 символов",
            "code": "WEAK_PASSWORD",
        }), 400

    with db() as con:
        duplicate = con.execute(
            "SELECT 1 FROM engineers "
            "WHERE lower(first_name) = lower(?) AND lower(name) = lower(?)",
            (first_name, last_name),
        ).fetchone()

        if duplicate:
            return jsonify({
                "error": (
                    f"Инженер «{first_name} {last_name}» уже "
                    "зарегистрирован. Если это другой человек, "
                    "добавьте цифру к фамилии, например "
                    f"«{last_name} 2»."
                ),
                "code": "DUPLICATE_NAME",
            }), 400

        key = generate_engineer_key(con, first_name, last_name)
        code = generate_engineer_code(con)

        con.execute(
            "INSERT INTO engineers "
            "(key, name, first_name, code, password_hash, is_admin, "
            "status, created_at) "
            "VALUES (?, ?, ?, ?, ?, 0, 'pending', ?)",
            (
                key,
                last_name,
                first_name,
                code,
                generate_password_hash(password),
                datetime.now().isoformat(timespec="seconds"),
            ),
        )

    return jsonify({
        "status": "pending",
        "message": "Заявка отправлена. Дождитесь подтверждения администратором.",
    })


@app.route("/api/admin/pending-engineers", methods=["GET"])
def api_admin_pending_engineers():
    if not require_admin():
        return jsonify({
            "error": "Недостаточно прав",
            "code": "FORBIDDEN",
        }), 403

    with db() as con:
        rows = con.execute(
            "SELECT key, name, first_name, code, created_at FROM engineers "
            "WHERE status = 'pending' ORDER BY created_at"
        ).fetchall()

    return jsonify([
        {
            "id": row["key"],
            "name": row["name"],
            "first_name": row["first_name"] or "",
            "code": row["code"],
            "created_at": row["created_at"],
        }
        for row in rows
    ])


@app.route("/api/admin/engineers/<key>/approve", methods=["POST"])
def api_admin_approve_engineer(key):
    if not require_admin():
        return jsonify({
            "error": "Недостаточно прав",
            "code": "FORBIDDEN",
        }), 403

    with db() as con:
        result = con.execute(
            "UPDATE engineers SET status = 'active' "
            "WHERE key = ? AND status = 'pending'",
            (key,),
        )

        if result.rowcount == 0:
            return jsonify({
                "error": "Заявка не найдена",
                "code": "NOT_FOUND",
            }), 404

    return jsonify({"status": "active"})


@app.route("/api/admin/engineers/<key>/reject", methods=["POST"])
def api_admin_reject_engineer(key):
    if not require_admin():
        return jsonify({
            "error": "Недостаточно прав",
            "code": "FORBIDDEN",
        }), 403

    with db() as con:
        result = con.execute(
            "DELETE FROM engineers WHERE key = ? AND status = 'pending'",
            (key,),
        )

        if result.rowcount == 0:
            return jsonify({
                "error": "Заявка не найдена",
                "code": "NOT_FOUND",
            }), 404

    return jsonify({"status": "rejected"})


@app.route("/api/admin/engineers", methods=["GET"])
def api_admin_engineers():
    if not require_admin():
        return jsonify({
            "error": "Недостаточно прав",
            "code": "FORBIDDEN",
        }), 403

    with db() as con:
        rows = con.execute(
            "SELECT key, name, first_name, code, is_admin FROM engineers "
            "WHERE status = 'active' ORDER BY name"
        ).fetchall()

    return jsonify([
        {
            "id": row["key"],
            "name": row["name"],
            "first_name": row["first_name"] or "",
            "code": row["code"],
            "is_admin": bool(row["is_admin"]),
        }
        for row in rows
    ])


@app.route("/api/admin/engineers/<key>", methods=["DELETE"])
def api_admin_delete_engineer(key):
    if not require_admin():
        return jsonify({
            "error": "Недостаточно прав",
            "code": "FORBIDDEN",
        }), 403

    if key == session.get("engineer_key"):
        return jsonify({
            "error": "Нельзя удалить свою учётную запись",
            "code": "CANNOT_DELETE_SELF",
        }), 400

    if key == PROTECTED_ENGINEER_KEY:
        return jsonify({
            "error": "Эту учётную запись нельзя удалить",
            "code": "PROTECTED_ENGINEER",
        }), 400

    with db() as con:
        row = con.execute(
            "SELECT is_admin FROM engineers WHERE key = ? AND status = 'active'",
            (key,),
        ).fetchone()

        if not row:
            return jsonify({
                "error": "Инженер не найден",
                "code": "NOT_FOUND",
            }), 404

        if row["is_admin"]:
            admin_count = con.execute(
                "SELECT COUNT(*) AS n FROM engineers "
                "WHERE is_admin = 1 AND status = 'active'"
            ).fetchone()["n"]

            if admin_count <= 1:
                return jsonify({
                    "error": "Нельзя удалить последнего администратора",
                    "code": "LAST_ADMIN",
                }), 400

        # Soft delete: keep the row so its key/code can never be
        # reassigned to someone else later (act numbering depends on
        # a code staying tied to one person forever).
        con.execute(
            "UPDATE engineers SET status = 'deleted' WHERE key = ?",
            (key,),
        )

    return jsonify({"status": "deleted"})


@app.route("/api/admin/engineers/<key>/promote", methods=["POST"])
def api_admin_promote_engineer(key):
    if not require_admin():
        return jsonify({
            "error": "Недостаточно прав",
            "code": "FORBIDDEN",
        }), 403

    with db() as con:
        result = con.execute(
            "UPDATE engineers SET is_admin = 1 "
            "WHERE key = ? AND status = 'active'",
            (key,),
        )

        if result.rowcount == 0:
            return jsonify({
                "error": "Инженер не найден",
                "code": "NOT_FOUND",
            }), 404

    return jsonify({"is_admin": True})


@app.route("/api/admin/engineers/<key>/demote", methods=["POST"])
def api_admin_demote_engineer(key):
    if not require_admin():
        return jsonify({
            "error": "Недостаточно прав",
            "code": "FORBIDDEN",
        }), 403

    if key == PROTECTED_ENGINEER_KEY:
        return jsonify({
            "error": "Нельзя снять права администратора с этой учётной записи",
            "code": "PROTECTED_ENGINEER",
        }), 400

    with db() as con:
        row = con.execute(
            "SELECT is_admin FROM engineers WHERE key = ? AND status = 'active'",
            (key,),
        ).fetchone()

        if not row:
            return jsonify({
                "error": "Инженер не найден",
                "code": "NOT_FOUND",
            }), 404

        if row["is_admin"]:
            admin_count = con.execute(
                "SELECT COUNT(*) AS n FROM engineers "
                "WHERE is_admin = 1 AND status = 'active'"
            ).fetchone()["n"]

            if admin_count <= 1:
                return jsonify({
                    "error": "Нельзя снять права с последнего администратора",
                    "code": "LAST_ADMIN",
                }), 400

        con.execute(
            "UPDATE engineers SET is_admin = 0 WHERE key = ?",
            (key,),
        )

    return jsonify({"is_admin": False})


@app.route("/api/intraservice/search")
def api_intraservice_search():
    if not get_engineer():
        return unauthenticated()

    query = clean(request.args.get("q", ""))
    if not query:
        return {"tasks": []}

    try:
        return {"tasks": search_intraservice(query)}
    except requests.HTTPError as exc:
        return {"error": f"Intraservice HTTP error: {exc}"}, 502
    except requests.RequestException as exc:
        return {"error": f"Intraservice connection error: {exc}"}, 502
    except Exception as exc:
        return {"error": f"Intraservice error: {exc}"}, 500


@app.route("/api/engineers", methods=["GET"])
def api_engineers():
    with db() as con:
        rows = con.execute(
            "SELECT key, name, first_name, code, password_hash "
            "FROM engineers WHERE status = 'active' ORDER BY name"
        ).fetchall()

    return jsonify([
        {
            "id": row["key"],
            "name": row["name"],
            "first_name": row["first_name"] or "",
            "code": row["code"],
            "has_password": row["password_hash"] is not None,
        }
        for row in rows
    ])


@app.route("/api/devices", methods=["GET"])
def api_devices():
    if not get_engineer():
        return jsonify({
            "error": "Не выбран инженер",
            "code": "UNAUTHENTICATED",
        }), 401

    q = clean(request.args.get("q", ""))

    if not q:
        return jsonify([])

    q_norm = normalize_serial(q)

    with db() as con:
        rows = con.execute(
            """
            SELECT
                id,
                customer_name,
                model,
                serial_number,
                address
            FROM devices
            WHERE serial_normalized LIKE ?
            ORDER BY customer_name, model
            LIMIT 100
            """,
            (f"%{q_norm}%",),
        ).fetchall()

    return jsonify([
        {
            "id": row["id"],
            "customer_name": row["customer_name"] or "",
            "device_model": row["model"] or "",
            "serial_number": row["serial_number"] or "",
            "address": row["address"] or "",
        }
        for row in rows
    ])


ACT_TEXT_FIELDS = [
    "customer_name",
    "customer_representative",
    "device_model",
    "serial_number",
    "printeco_label",
    "device_type",
    "comment_label",
    "address",
    "phone",
    "fault",
    "counter_bw",
    "counter_color",
    "service_kind",
    "diagnostics_result",
    "works_text",
    "work_date",
    "start_time",
    "end_time",
    "customer_signatory",
    "intraservice_task_id",
]

# Free-text fields where the engineer may use several lines.
ACT_MULTILINE_FIELDS = {"fault", "diagnostics_result", "works_text"}

ACT_REQUIRED_FIELDS = {
    "customer_name": "Укажите клиента",
    "serial_number": "Укажите серийный номер",
    "work_date": "Укажите дату работ",
}

# A thermo act (ремонт узла терморегистрации) only has a customer, a date
# and spare parts; every other text field stays empty.
THERMO_FIELDS = {"customer_name", "work_date"}

ACT_LIST_SQL = """
    SELECT
        a.*,
        repaired.act_number AS repaired_by_act_number,
        original.act_number AS repair_of_act_number,
        deleter.name AS deleted_by_name
    FROM acts a
    LEFT JOIN acts repaired ON repaired.id = a.repaired_by_act_id
    LEFT JOIN acts original ON original.id = a.repair_of_act_id
    LEFT JOIN engineers deleter ON deleter.key = a.deleted_by
"""


def api_error(message, code, http_status):
    return jsonify({"error": message, "code": code}), http_status


def unauthenticated():
    return api_error("Не выбран инженер", "UNAUTHENTICATED", 401)


def act_state(act):
    """working / broken / repaired — what the UI shows as the act badge.

    Thermo acts have no condition, so they have no state either.
    """
    if act["number_type"] == "thermo":
        return None
    if act["repaired_by_act_id"]:
        return "repaired"
    if act["device_condition"] == CONDITION_BROKEN:
        return "broken"
    return "working"


def can_edit_act(act, engineer_key, engineer):
    if act["deleted_at"]:
        return False
    if engineer["is_admin"]:
        return True
    # A repaired original is a historical record: admins only.
    if act["repaired_by_act_id"]:
        return False
    return act["engineer_key"] == engineer_key


def can_repair_act(act):
    return (
        not act["deleted_at"]
        and act["device_condition"] == CONDITION_BROKEN
        and not act["repaired_by_act_id"]
    )


def can_delete_act(act, engineer_key, engineer):
    if act["deleted_at"]:
        return False
    # An original with a live repaired copy: the copy must be deleted first.
    if act["repaired_by_act_id"]:
        return False
    return engineer["is_admin"] or act["engineer_key"] == engineer_key


def linked_act(act_id, act_number):
    if not act_id:
        return None
    return {"id": act_id, "act_number": act_number or ""}


def act_list_item(act):
    return {
        "id": act["id"],
        "act_number": act["act_number"] or "",
        "number_type": act["number_type"] or "standard",
        "device_condition": act["device_condition"] or "",
        "state": act_state(act),
        "customer_name": act["customer_name"] or "",
        "device_model": act["device_model"] or "",
        "serial_number": act["serial_number"] or "",
        "engineer_name": act["engineer_name"] or "",
        "engineer_key": act["engineer_key"] or "",
        "created_at": act["created_at"] or "",
        "work_date": act["work_date"] or "",
        "repaired_by": linked_act(
            act["repaired_by_act_id"], act["repaired_by_act_number"]
        ),
        "repair_of": linked_act(
            act["repair_of_act_id"], act["repair_of_act_number"]
        ),
    }


def load_act(con, act_id):
    return con.execute(
        ACT_LIST_SQL + " WHERE a.id = ?",
        (act_id,),
    ).fetchone()


def act_detail(con, act, engineer_key, engineer):
    data = act_list_item(act)
    data.update({
        field_name: act[field_name] or "" for field_name in ACT_TEXT_FIELDS
    })
    data.update({
        "engineer_code": act["engineer_code"] or "",
        "source_device_id": act["source_device_id"],
        "updated_at": act["updated_at"] or "",
        "materials": get_act_materials(con, act["id"]),
        "can_edit": can_edit_act(act, engineer_key, engineer),
        "can_repair": can_repair_act(act),
        "can_delete": can_delete_act(act, engineer_key, engineer),
        "is_deleted": bool(act["deleted_at"]),
        "deleted_at": act["deleted_at"] or "",
        "deleted_by_name": act["deleted_by_name"] or "",
        "original_act_number": act["original_act_number"] or "",
    })
    return data


def read_act_payload(data, number_type="standard"):
    """Validate the JSON body of create/update/repair requests.

    Returns (values, materials, error_response). error_response is None when
    the payload is valid.
    """
    if not isinstance(data, dict):
        return None, None, api_error(
            "Неверный формат запроса", "BAD_REQUEST", 400
        )

    thermo = number_type == "thermo"
    values = {}
    for name in ACT_TEXT_FIELDS:
        raw = data.get(name, "")
        if thermo and name not in THERMO_FIELDS:
            raw = ""
        if not isinstance(raw, str):
            raw = "" if raw is None else str(raw)
        values[name] = (
            clean_multiline(raw)
            if name in ACT_MULTILINE_FIELDS
            else clean(raw)
        )

    for name, message in ACT_REQUIRED_FIELDS.items():
        if thermo and name not in THERMO_FIELDS:
            continue
        if not values[name]:
            return None, None, api_error(message, "VALIDATION_ERROR", 400)

    try:
        datetime.strptime(values["work_date"], "%Y-%m-%d")
    except ValueError:
        return None, None, api_error(
            "Неверный формат даты работ", "VALIDATION_ERROR", 400
        )

    raw_materials = data.get("materials", [])
    if not isinstance(raw_materials, list) or not all(
        isinstance(item, dict) for item in raw_materials
    ):
        return None, None, api_error(
            "Неверный формат материалов", "VALIDATION_ERROR", 400
        )

    materials = []
    for item in raw_materials:
        material = {
            key: clean(str(item.get(key) or ""))
            for key in ("name", "article", "quantity")
        }
        if material["name"] or material["article"] or material["quantity"]:
            materials.append(material)

    return values, materials, None


def resolve_source_device(con, data, values, now):
    """Use the device picked from the catalogue, or remember a manual one."""
    device_id = data.get("source_device_id")
    if isinstance(device_id, int):
        exists = con.execute(
            "SELECT 1 FROM devices WHERE id = ?",
            (device_id,),
        ).fetchone()
        if exists:
            return device_id
    return save_manual_device(con, values, now)


def insert_act(
    con,
    engineer_key,
    engineer,
    values,
    materials,
    condition,
    number_type,
    source_device_id,
    now,
    repair_of_act_id=None,
):
    number = next_act_number(
        con, engineer["code"], now, condition, number_type
    )
    timestamp = now.isoformat(timespec="seconds")
    columns = [
        "act_number", "number_type", "engineer_key", "engineer_name",
        "engineer_code", "created_at", "updated_at", "status",
        "device_condition", "source_device_id",
        "repair_of_act_id",
    ] + ACT_TEXT_FIELDS
    row = [
        number, number_type, engineer_key, engineer["name"],
        engineer["code"], timestamp, timestamp, "completed",
        condition, source_device_id,
        repair_of_act_id,
    ] + [values[name] for name in ACT_TEXT_FIELDS]

    cursor = con.execute(
        f"INSERT INTO acts ({', '.join(columns)}) "
        f"VALUES ({', '.join('?' for _ in columns)})",
        row,
    )
    save_act_materials(con, cursor.lastrowid, materials)
    return cursor.lastrowid


@app.route("/api/acts", methods=["GET"])
def api_acts_list():
    if not get_engineer():
        return unauthenticated()

    number_type = clean(request.args.get("type", "")) or "standard"
    if number_type not in NUMBER_TYPES:
        number_type = "standard"
    thermo = number_type == "thermo"

    # The thermo list only has customer and date filters.
    q = "" if thermo else clean(request.args.get("q", ""))
    customer = clean(request.args.get("customer", "")) if thermo else ""
    engineer_filter = (
        "" if thermo else clean(request.args.get("engineer", ""))
    )
    state_filter = "" if thermo else clean(request.args.get("state", ""))
    date_from = clean(request.args.get("date_from", ""))
    date_to = clean(request.args.get("date_to", ""))

    sql = ACT_LIST_SQL + " WHERE a.number_type = ? AND a.deleted_at IS NULL"
    params = [number_type]

    if engineer_filter in get_engineers_dict():
        sql += " AND a.engineer_key = ?"
        params.append(engineer_filter)

    # "broken" includes repaired originals: their condition stays broken,
    # only the badge in the list says "repaired".
    if state_filter == "working":
        sql += " AND a.device_condition = ?"
        params.append(CONDITION_WORKING)
    elif state_filter == "broken":
        sql += " AND a.device_condition = ?"
        params.append(CONDITION_BROKEN)

    if date_from:
        sql += " AND DATE(a.created_at) >= DATE(?)"
        params.append(date_from)

    if date_to:
        sql += " AND DATE(a.created_at) <= DATE(?)"
        params.append(date_to)

    sql += " ORDER BY datetime(a.created_at) DESC, a.id DESC"

    with db() as con:
        rows = con.execute(sql, params).fetchall()

    if q:
        # SQLite's LIKE is only case-insensitive for ASCII, so Cyrillic
        # searches ("дикси" vs "ДИКСИ") need Python-side filtering.
        q_norm = q.lower()
        rows = [
            row
            for row in rows
            if (
                q_norm in (row["act_number"] or "").lower()
                or q_norm in (row["serial_number"] or "").lower()
                or q_norm in (row["customer_name"] or "").lower()
                or q_norm in (row["device_model"] or "").lower()
            )
        ]

    if customer:
        customer_norm = customer.lower()
        rows = [
            row
            for row in rows
            if customer_norm in (row["customer_name"] or "").lower()
        ]

    rows = rows[:200]

    return jsonify([act_list_item(row) for row in rows])


@app.route("/api/acts/<int:act_id>", methods=["GET"])
def api_act_detail(act_id):
    engineer = get_engineer()
    if not engineer:
        return unauthenticated()

    with db() as con:
        act = load_act(con, act_id)
        # Deleted acts are visible only to admins (the "Удалённые акты" page).
        if not act or (act["deleted_at"] and not engineer["is_admin"]):
            return api_error("Акт не найден", "NOT_FOUND", 404)
        return jsonify(
            act_detail(con, act, session["engineer_key"], engineer)
        )


@app.route("/api/acts", methods=["POST"])
def api_act_create():
    engineer = get_engineer()
    if not engineer:
        return unauthenticated()

    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return api_error("Неверный формат запроса", "BAD_REQUEST", 400)

    number_type = data.get("number_type") or "standard"
    if number_type not in NUMBER_TYPES:
        return api_error("Неверный тип акта", "VALIDATION_ERROR", 400)

    values, materials, error = read_act_payload(data, number_type)
    if error:
        return error

    if number_type == "thermo":
        condition = ""
    else:
        condition = clean(data.get("device_condition"))
        if condition not in DEVICE_CONDITIONS:
            return api_error(
                "Укажите состояние аппарата", "VALIDATION_ERROR", 400
            )

    engineer_key = session["engineer_key"]
    now = datetime.now()

    with db() as con:
        source_device_id = (
            None if number_type == "thermo"
            else resolve_source_device(con, data, values, now)
        )
        act_id = insert_act(
            con, engineer_key, engineer, values, materials,
            condition, number_type, source_device_id, now,
        )
        act = load_act(con, act_id)
        return jsonify(act_detail(con, act, engineer_key, engineer)), 201


@app.route("/api/acts/<int:act_id>", methods=["PUT"])
def api_act_update(act_id):
    engineer = get_engineer()
    if not engineer:
        return unauthenticated()

    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return api_error("Неверный формат запроса", "BAD_REQUEST", 400)

    engineer_key = session["engineer_key"]
    now = datetime.now()

    with db() as con:
        act = load_act(con, act_id)
        if not act or act["deleted_at"]:
            return api_error("Акт не найден", "NOT_FOUND", 404)

        if not can_edit_act(act, engineer_key, engineer):
            return api_error(
                "Нет прав на редактирование этого акта", "FORBIDDEN", 403
            )

        thermo = act["number_type"] == "thermo"
        values, materials, error = read_act_payload(data, act["number_type"])
        if error:
            return error

        # Condition and number type define the act number: frozen forever.
        condition = data.get("device_condition")
        if (
            not thermo
            and condition is not None
            and clean(condition) != act["device_condition"]
        ):
            return api_error(
                "Состояние аппарата нельзя изменить после сохранения",
                "CONDITION_FROZEN",
                400,
            )
        number_type = data.get("number_type")
        if number_type is not None and number_type != act["number_type"]:
            return api_error(
                "Тип акта нельзя изменить после сохранения",
                "NUMBER_TYPE_FROZEN",
                400,
            )

        # Only an admin may change the act number by hand (free format).
        act_number = act["act_number"]
        if "act_number" in data:
            act_number = clean(str(data.get("act_number") or ""))
        if act_number != act["act_number"]:
            if not engineer["is_admin"]:
                return api_error(
                    "Изменить номер акта может только администратор",
                    "FORBIDDEN",
                    403,
                )
            if not act_number:
                return api_error(
                    "Номер акта не может быть пустым", "VALIDATION_ERROR", 400
                )
            # DEL-NNNN is reserved for deleted acts.
            if act_number.upper().startswith("DEL-"):
                return api_error(
                    "Номера DEL-… зарезервированы для удалённых актов",
                    "VALIDATION_ERROR",
                    400,
                )
            duplicate = con.execute(
                "SELECT 1 FROM acts WHERE act_number = ? AND id <> ?",
                (act_number, act_id),
            ).fetchone()
            if duplicate:
                return api_error(
                    f"Номер {act_number} уже занят другим актом",
                    "NUMBER_TAKEN",
                    409,
                )

        # Keep the previous state as a version, but only if something changed.
        old_materials = get_act_materials(con, act_id)
        changed = (
            act_number != act["act_number"]
            or materials != old_materials
            or any(
                values[name] != (act[name] or "") for name in ACT_TEXT_FIELDS
            )
        )
        if changed:
            save_act_version(con, act, engineer_key, engineer, now)

        source_device_id = (
            None if thermo
            else resolve_source_device(con, data, values, now)
        )
        assignments =", ".join(f"{name} = ?" for name in ACT_TEXT_FIELDS)
        con.execute(
            f"UPDATE acts SET {assignments}, act_number = ?, "
            "source_device_id = ?, updated_at = ? WHERE id = ?",
            [values[name] for name in ACT_TEXT_FIELDS] + [
                act_number,
                source_device_id,
                now.isoformat(timespec="seconds"),
                act_id,
            ],
        )
        save_act_materials(con, act_id, materials)

        act = load_act(con, act_id)
        return jsonify(act_detail(con, act, engineer_key, engineer))


@app.route("/api/acts/<int:act_id>/repair", methods=["POST"])
def api_act_repair(act_id):
    """Mark a broken act as repaired by creating a working copy of it."""
    engineer = get_engineer()
    if not engineer:
        return unauthenticated()

    data = request.get_json(silent=True)
    values, materials, error = read_act_payload(data)
    if error:
        return error

    engineer_key = session["engineer_key"]
    now = datetime.now()

    with db() as con:
        original = load_act(con, act_id)
        if not original or original["deleted_at"]:
            return api_error("Акт не найден", "NOT_FOUND", 404)

        if original["device_condition"] != CONDITION_BROKEN:
            return api_error(
                "Отремонтировать можно только акт «Не работает»",
                "NOT_BROKEN",
                400,
            )
        if original["repaired_by_act_id"]:
            return api_error(
                "Этот акт уже отремонтирован",
                "ALREADY_REPAIRED",
                409,
            )

        source_device_id = resolve_source_device(con, data, values, now)
        copy_id = insert_act(
            con, engineer_key, engineer, values, materials,
            CONDITION_WORKING, original["number_type"] or "standard",
            source_device_id, now, repair_of_act_id=act_id,
        )
        con.execute(
            "UPDATE acts SET repaired_by_act_id = ?, updated_at = ? "
            "WHERE id = ?",
            (copy_id, now.isoformat(timespec="seconds"), act_id),
        )

        copy = load_act(con, copy_id)
        return jsonify(act_detail(con, copy, engineer_key, engineer)), 201


@app.route("/api/acts/stats", methods=["GET"])
def api_acts_stats():
    """Counters for the home page: all time and the current calendar month.

    The month is taken from work_date. "broken" includes repaired originals,
    the same way the list filter does.
    """
    if not get_engineer():
        return unauthenticated()

    month_prefix = datetime.now().strftime("%Y-%m-") + "%"
    sql = """
        SELECT
            SUM(number_type = 'standard' AND device_condition = ?) AS working,
            SUM(number_type = 'standard' AND device_condition = ?) AS broken,
            SUM(number_type = 'thermo') AS thermo
        FROM acts
        WHERE deleted_at IS NULL
    """

    with db() as con:
        total = con.execute(
            sql, (CONDITION_WORKING, CONDITION_BROKEN)
        ).fetchone()
        month = con.execute(
            sql + " AND work_date LIKE ?",
            (CONDITION_WORKING, CONDITION_BROKEN, month_prefix),
        ).fetchone()

    def counters(row):
        return {key: row[key] or 0 for key in ("working", "broken", "thermo")}

    return jsonify({"total": counters(total), "month": counters(month)})


def next_sequence_value(con, name):
    """Next value of a counter that never repeats, even after purges."""
    con.execute(
        "INSERT INTO sequences (name, value) VALUES (?, 1) "
        "ON CONFLICT(name) DO UPDATE SET value = value + 1",
        (name,),
    )
    return con.execute(
        "SELECT value FROM sequences WHERE name = ?",
        (name,),
    ).fetchone()["value"]


@app.route("/api/acts/<int:act_id>", methods=["DELETE"])
def api_act_delete(act_id):
    """Move an act to the admin-only "Удалённые акты" list.

    The act gets a DEL-NNNN number; its old number is kept in
    original_act_number and may be given to a new act later.
    """
    engineer = get_engineer()
    if not engineer:
        return unauthenticated()

    engineer_key = session["engineer_key"]
    now = datetime.now().isoformat(timespec="seconds")

    with db() as con:
        act = load_act(con, act_id)
        if not act or act["deleted_at"]:
            return api_error("Акт не найден", "NOT_FOUND", 404)

        if act["repaired_by_act_id"]:
            return api_error(
                f"Сначала удалите акт ремонта {act['repaired_by_act_number']}",
                "HAS_REPAIR_COPY",
                409,
            )
        if not can_delete_act(act, engineer_key, engineer):
            return api_error(
                "Нет прав на удаление этого акта", "FORBIDDEN", 403
            )

        deleted_number = f"DEL-{next_sequence_value(con, 'deleted_acts'):04d}"
        con.execute(
            "UPDATE acts SET act_number = ?, original_act_number = ?, "
            "deleted_at = ?, deleted_by = ?, updated_at = ? WHERE id = ?",
            (deleted_number, act["act_number"], now, engineer_key, now, act_id),
        )
        # Deleting a repaired copy makes the original "not repaired" again.
        if act["repair_of_act_id"]:
            con.execute(
                "UPDATE acts SET repaired_by_act_id = NULL, updated_at = ? "
                "WHERE id = ? AND repaired_by_act_id = ?",
                (now, act["repair_of_act_id"], act_id),
            )

    return jsonify({"status": "deleted", "act_number": deleted_number})


@app.route("/api/admin/deleted-acts", methods=["GET"])
def api_admin_deleted_acts():
    if not require_admin():
        return api_error("Нет прав", "FORBIDDEN", 403)

    with db() as con:
        rows = con.execute(
            ACT_LIST_SQL
            + " WHERE a.deleted_at IS NOT NULL"
            " ORDER BY a.deleted_at DESC, a.id DESC"
        ).fetchall()

    return jsonify([
        {
            "id": row["id"],
            "act_number": row["act_number"] or "",
            "original_act_number": row["original_act_number"] or "",
            "number_type": row["number_type"] or "standard",
            "customer_name": row["customer_name"] or "",
            "engineer_name": row["engineer_name"] or "",
            "deleted_at": row["deleted_at"] or "",
            "deleted_by_name": row["deleted_by_name"] or "",
        }
        for row in rows
    ])


def restore_deleted_act(con, act, now):
    """Bring a deleted act back with a new number (inside the caller's
    transaction).

    The number uses the author's engineer code and the current month.
    A restored repaired copy is linked to its original again, unless the
    original has been repaired by another copy in the meantime.
    Returns (new_number, None) or (None, error_response).
    """
    timestamp = now.isoformat(timespec="seconds")

    original = (
        load_act(con, act["repair_of_act_id"])
        if act["repair_of_act_id"]
        else None
    )
    relink = False
    if original and not original["deleted_at"]:
        if original["repaired_by_act_id"]:
            return None, (jsonify({
                "error": (
                    f"Исходный акт {original['act_number']} уже "
                    f"отремонтирован другим актом "
                    f"{original['repaired_by_act_number']}"
                ),
                "code": "ORIGINAL_REPAIRED",
                "repaired_by": {
                    "id": original["repaired_by_act_id"],
                    "act_number": original["repaired_by_act_number"],
                },
            }), 409)
        relink = True

    number = next_act_number(
        con,
        act["engineer_code"],
        now,
        act["device_condition"],
        act["number_type"] or "standard",
    )
    con.execute(
        "UPDATE acts SET act_number = ?, deleted_at = NULL, "
        "deleted_by = NULL, original_act_number = NULL, updated_at = ? "
        "WHERE id = ?",
        (number, timestamp, act["id"]),
    )
    if relink:
        con.execute(
            "UPDATE acts SET repaired_by_act_id = ?, updated_at = ? "
            "WHERE id = ?",
            (act["id"], timestamp, original["id"]),
        )
    return number, None


@app.route("/api/admin/deleted-acts/<int:act_id>/restore", methods=["POST"])
def api_admin_restore_deleted_act(act_id):
    if not require_admin():
        return api_error("Нет прав", "FORBIDDEN", 403)

    with db() as con:
        act = load_act(con, act_id)
        if not act or not act["deleted_at"]:
            return api_error("Удалённый акт не найден", "NOT_FOUND", 404)

        number, error = restore_deleted_act(con, act, datetime.now())
        if error:
            return error

    return jsonify({"id": act_id, "act_number": number})


PURGE_PERIODS = {"all": None, "month": 30, "week": 7}


@app.route("/api/admin/deleted-acts/purge", methods=["POST"])
def api_admin_purge_deleted_acts():
    """Permanently remove deleted acts.

    Body: {"ids": [...]} for selected acts, or {"period": "all"|"month"|"week"}
    for everything deleted within the last 30 / 7 days.
    """
    if not require_admin():
        return api_error("Нет прав", "FORBIDDEN", 403)

    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return api_error("Неверный формат запроса", "BAD_REQUEST", 400)

    sql = "SELECT id FROM acts WHERE deleted_at IS NOT NULL"
    params = []

    if "ids" in data:
        ids = data["ids"]
        if not isinstance(ids, list) or not all(
            isinstance(item, int) for item in ids
        ):
            return api_error("Неверный список актов", "BAD_REQUEST", 400)
        if not ids:
            return jsonify({"purged": 0})
        sql += f" AND id IN ({', '.join('?' for _ in ids)})"
        params.extend(ids)
    elif data.get("period") in PURGE_PERIODS:
        days = PURGE_PERIODS[data["period"]]
        if days is not None:
            since = datetime.now() - timedelta(days=days)
            sql += " AND deleted_at >= ?"
            params.append(since.isoformat(timespec="seconds"))
    else:
        return api_error("Не выбраны акты для очистки", "BAD_REQUEST", 400)

    with db() as con:
        ids = [row["id"] for row in con.execute(sql, params).fetchall()]
        if ids:
            marks = ", ".join("?" for _ in ids)
            # Other acts may still point at the purged ones.
            con.execute(
                f"UPDATE acts SET repair_of_act_id = NULL "
                f"WHERE repair_of_act_id IN ({marks})",
                ids,
            )
            con.execute(
                f"UPDATE acts SET repaired_by_act_id = NULL "
                f"WHERE repaired_by_act_id IN ({marks})",
                ids,
            )
            con.execute(
                f"DELETE FROM act_materials WHERE act_id IN ({marks})", ids
            )
            con.execute(
                f"DELETE FROM act_versions WHERE act_id IN ({marks})", ids
            )
            con.execute(f"DELETE FROM acts WHERE id IN ({marks})", ids)

    return jsonify({"purged": len(ids)})


def save_act_version(con, act, engineer_key, engineer, now):
    """Store the current (pre-edit) state of an act as <number>/N."""
    index = next_sequence_value(con, f"act_versions:{act['id']}")
    snapshot = act_detail(con, act, engineer_key, engineer)
    con.execute(
        "INSERT INTO act_versions "
        "(act_id, version_number, created_at, created_by, data) "
        "VALUES (?, ?, ?, ?, ?)",
        (
            act["id"],
            f"{act['act_number']}/{index}",
            now.isoformat(timespec="seconds"),
            engineer_key,
            json.dumps(snapshot, ensure_ascii=False),
        ),
    )


ACT_VERSIONS_SQL = """
    SELECT
        v.id, v.act_id, v.version_number, v.created_at, v.data,
        editor.name AS created_by_name,
        a.act_number AS current_act_number,
        a.deleted_at AS act_deleted_at
    FROM act_versions v
    JOIN acts a ON a.id = v.act_id
    LEFT JOIN engineers editor ON editor.key = v.created_by
"""


def act_version_item(row):
    data = json.loads(row["data"])
    return {
        "id": row["id"],
        "act_id": row["act_id"],
        "version_number": row["version_number"],
        "created_at": row["created_at"],
        "created_by_name": row["created_by_name"] or "",
        "current_act_number": row["current_act_number"] or "",
        "act_deleted": bool(row["act_deleted_at"]),
        "number_type": data.get("number_type", "standard"),
        "customer_name": data.get("customer_name", ""),
    }


@app.route("/api/admin/act-versions", methods=["GET"])
def api_admin_act_versions():
    if not require_admin():
        return api_error("Нет прав", "FORBIDDEN", 403)

    with db() as con:
        rows = con.execute(
            ACT_VERSIONS_SQL + " ORDER BY v.created_at DESC, v.id DESC"
        ).fetchall()

    return jsonify([act_version_item(row) for row in rows])


@app.route("/api/admin/act-versions/<int:version_id>", methods=["GET"])
def api_admin_act_version(version_id):
    if not require_admin():
        return api_error("Нет прав", "FORBIDDEN", 403)

    with db() as con:
        row = con.execute(
            ACT_VERSIONS_SQL + " WHERE v.id = ?",
            (version_id,),
        ).fetchone()

    if not row:
        return api_error("Версия не найдена", "NOT_FOUND", 404)

    item = act_version_item(row)
    item["act"] = json.loads(row["data"])
    return jsonify(item)


@app.route(
    "/api/admin/act-versions/<int:version_id>/restore", methods=["POST"]
)
def api_admin_restore_act_version(version_id):
    """Replace the act's content with a version and drop all its versions.

    The act number stays as it is now; the current content is not kept.
    If the act is deleted, it is restored first (new number), all in one
    transaction: either both steps happen or neither.
    """
    engineer = get_engineer()
    if not engineer or not engineer["is_admin"]:
        return api_error("Нет прав", "FORBIDDEN", 403)

    now = datetime.now()

    with db() as con:
        version = con.execute(
            "SELECT act_id, data FROM act_versions WHERE id = ?",
            (version_id,),
        ).fetchone()
        if not version:
            return api_error("Версия не найдена", "NOT_FOUND", 404)

        act = load_act(con, version["act_id"])
        if act["deleted_at"]:
            _, error = restore_deleted_act(con, act, now)
            if error:
                # restore_deleted_act checks before writing, so nothing has
                # changed; rollback is just a safety net.
                con.rollback()
                return error

        snapshot = json.loads(version["data"])
        assignments = ", ".join(f"{name} = ?" for name in ACT_TEXT_FIELDS)
        con.execute(
            f"UPDATE acts SET {assignments}, source_device_id = ?, "
            "updated_at = ? WHERE id = ?",
            [snapshot.get(name, "") for name in ACT_TEXT_FIELDS] + [
                snapshot.get("source_device_id"),
                now.isoformat(timespec="seconds"),
                act["id"],
            ],
        )
        save_act_materials(con, act["id"], snapshot.get("materials", []))
        con.execute(
            "DELETE FROM act_versions WHERE act_id = ?", (act["id"],)
        )

        act = load_act(con, act["id"])
        return jsonify(
            act_detail(con, act, session["engineer_key"], engineer)
        )


@app.route("/api/admin/act-versions/purge", methods=["POST"])
def api_admin_purge_act_versions():
    """Permanently remove versions: {"ids": [...]} or {"period": ...}."""
    if not require_admin():
        return api_error("Нет прав", "FORBIDDEN", 403)

    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return api_error("Неверный формат запроса", "BAD_REQUEST", 400)

    sql = "DELETE FROM act_versions WHERE 1=1"
    params = []

    if "ids" in data:
        ids = data["ids"]
        if not isinstance(ids, list) or not all(
            isinstance(item, int) for item in ids
        ):
            return api_error("Неверный список версий", "BAD_REQUEST", 400)
        if not ids:
            return jsonify({"purged": 0})
        sql += f" AND id IN ({', '.join('?' for _ in ids)})"
        params.extend(ids)
    elif data.get("period") in PURGE_PERIODS:
        days = PURGE_PERIODS[data["period"]]
        if days is not None:
            since = datetime.now() - timedelta(days=days)
            sql += " AND created_at >= ?"
            params.append(since.isoformat(timespec="seconds"))
    else:
        return api_error("Не выбраны версии для очистки", "BAD_REQUEST", 400)

    with db() as con:
        purged = con.execute(sql, params).rowcount

    return jsonify({"purged": purged})


@app.route("/api/admin/devices/sync", methods=["POST"])
def api_admin_sync_devices():
    """Reload the device catalogue from the Google Sheet."""
    if not require_admin():
        return api_error("Нет прав", "FORBIDDEN", 403)

    try:
        count = import_google_sheet()
    except Exception as exc:
        return api_error(
            f"Не удалось загрузить Google-таблицу: {exc}", "SYNC_FAILED", 502
        )
    return jsonify({"imported": count})


# Run schema migrations on import, so they are applied no matter how the app
# is started (waitress imports app:app and never reaches __main__).
# init_db is idempotent: every step checks what already exists.
init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
