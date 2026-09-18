import csv
import io
import os
import sqlite3
import urllib.request
from contextlib import contextmanager
from datetime import datetime
from io import BytesIO

import requests
from dotenv import load_dotenv
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

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

ENGINEERS = {
    "podkolodny": {"name": "Подколодный", "code": "1"},
    "dzyuba": {"name": "Дзюба", "code": "2"},
    "izyurov": {"name": "Изъюров", "code": "5"},
    "ozerov": {"name": "Озеров", "code": "9"},
}

NUMBER_TYPES = {
    "standard": {"name": "Обычный акт"},
    "thermo": {"name": "Узел терморегистрации"},
}

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get(
    "FLASK_SECRET_KEY",
    "change-this-before-network-use",
)

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

def get_engineer():
    return ENGINEERS.get(session.get("engineer_key"))

def require_engineer():
    if not get_engineer():
        flash("Сначала выберите инженера.", "warning")
        return False
    return True

def next_act_number(
    con,
    engineer_code,
    now,
    status="completed",
    number_type="standard",
):
    month = now.strftime("%m")

    if status == "draft" and number_type == "thermo":
        prefix = f"DT-{engineer_code}-{month}-"
    elif status == "draft":
        prefix = f"D-{engineer_code}-{month}-"
    elif number_type == "thermo":
        prefix = f"T-{engineer_code}-{month}-"
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

@app.route("/")
def index():
    if get_engineer():
        return redirect(url_for("acts_list"))
    return render_template("engineer.html")

@app.route("/engineer", methods=["POST"])
def engineer_select():
    key = request.form.get("engineer_key")
    if key in ENGINEERS:
        session["engineer_key"] = key
        session.permanent = True
        return redirect(url_for("devices_search"))
    flash("Выберите инженера из списка.", "warning")
    return redirect(url_for("index"))

@app.route("/devices")
def devices_search():
    query = request.args.get("q", "")
    intraservice_results = []
    manual_devices = []
    google_devices = []

    if query:
        intraservice_results = search_intraservice(query)

        with db() as con:
            normalized = normalize_serial(query)
            manual_devices = con.execute(
                """
                SELECT id, customer_name, model, serial_number, address
                FROM devices
                WHERE source = 'manual'
                  AND serial_normalized LIKE ?
                ORDER BY customer_name, model
                LIMIT 50
                """,
                (f"%{normalized}%",),
            ).fetchall()

            google_devices = con.execute(
                """
                SELECT id, customer_name, model, serial_number, address
                FROM devices
                WHERE source = 'google_sheet_main'
                  AND serial_normalized LIKE ?
                ORDER BY customer_name, model
                LIMIT 50
                """,
                (f"%{normalized}%",),
            ).fetchall()

    return render_template(
        "devices.html",
        query=query,
        intraservice_results=intraservice_results,
        manual_devices=manual_devices,
        google_devices=google_devices,
    )

@app.route("/devices/save", methods=["POST"])
def device_save():
    if not require_engineer():
        return redirect(url_for("index"))

    now = datetime.now()
    values = {
        "serial_number": request.form.get("serial_number"),
        "customer_name": request.form.get("customer_name"),
        "address": request.form.get("address"),
        "device_model": request.form.get("device_model"),
    }

    with db() as con:
        device_id = save_manual_device(con, values, now)

    if device_id:
        flash("Аппарат сохранён.", "success")
    else:
        flash("Не удалось сохранить аппарат. Проверьте серийный номер и клиента.", "danger")

    return redirect(url_for("devices_search"))

@app.route("/acts")
def acts_list():
    if not require_engineer():
        return redirect(url_for("index"))

    q = request.args.get("q", "")
    engineer_filter = request.args.get("engineer", "")
    status_filter = request.args.get("status", "")
    date_sort = request.args.get("date_sort", "")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")

    with db() as con:
        query = """
            SELECT
                id, act_number, status, customer_name, device_model,
                serial_number, engineer_name, created_at
            FROM acts
            WHERE 1=1
        """
        params = []

        if q:
            query += " AND (act_number LIKE ? OR serial_number LIKE ? OR customer_name LIKE ?)"
            params.extend([f"%{q}%"] * 3)

        if engineer_filter:
            query += " AND engineer_key = ?"
            params.append(engineer_filter)

        if status_filter:
            query += " AND status = ?"
            params.append(status_filter)

        if date_from:
            query += " AND DATE(created_at) >= ?"
            params.append(date_from)

        if date_to:
            query += " AND DATE(created_at) <= ?"
            params.append(date_to)

        if date_sort == "asc":
            query += " ORDER BY created_at ASC"
        elif date_sort == "desc":
            query += " ORDER BY created_at DESC"
        else:
            query += " ORDER BY created_at DESC"

        acts = con.execute(query, params).fetchall()

    return render_template(
        "acts.html",
        acts=acts,
        q=q,
        engineer_filter=engineer_filter,
        status_filter=status_filter,
        date_sort=date_sort,
        date_from=date_from,
        date_to=date_to,
    )

@app.route("/acts/new")
def act_new():
    if not require_engineer():
        return redirect(url_for("index"))

    device_id = request.args.get("device_id")
    device = None
    intraservice_task = None

    if device_id:
        with db() as con:
            device = con.execute(
                "SELECT * FROM devices WHERE id = ?", (device_id,)
            ).fetchone()

        if device and device.get("intraservice_task_id"):
            intraservice_task = search_intraservice(device["intraservice_task_id"])
            intraservice_task = intraservice_task[0] if intraservice_task else None

    return render_template(
        "act_form.html",
        device=device,
        intraservice_task=intraservice_task,
        act=None,
        number_types=NUMBER_TYPES,
    )

@app.route("/acts/<int:act_id>")
def act_view(act_id):
    if not require_engineer():
        return redirect(url_for("index"))

    with db() as con:
        act = con.execute(
            "SELECT * FROM acts WHERE id = ?", (act_id,)
        ).fetchone()

    if not act:
        flash("Акт не найден.", "danger")
        return redirect(url_for("acts_list"))

    return render_template(
        "act_print.html",
        act=act,
        materials=parse_materials(act["materials_text"]),
    )

@app.route("/acts/<int:act_id>/edit")
def act_edit(act_id):
    if not require_engineer():
        return redirect(url_for("index"))

    with db() as con:
        act = con.execute(
            "SELECT * FROM acts WHERE id = ?", (act_id,)
        ).fetchone()

    if not act:
        flash("Акт не найден.", "danger")
        return redirect(url_for("acts_list"))

    return render_template(
        "act_form.html",
        device=None,
        intraservice_task=None,
        act=act,
        number_types=NUMBER_TYPES,
    )

@app.route("/acts/save", methods=["POST"])
def act_save():
    if not require_engineer():
        return redirect(url_for("index"))

    engineer = get_engineer()
    now = datetime.now()

    act_id = request.form.get("act_id")
    is_draft = request.form.get("is_draft") == "on"
    number_type = request.form.get("number_type", "standard")

    device_id = request.form.get("source_device_id")
    if device_id:
        device_id = int(device_id) if device_id else None

    with db() as con:
        if act_id:
            act = con.execute(
                "SELECT * FROM acts WHERE id = ?", (act_id,)
            ).fetchone()

            if not act:
                flash("Акт не найден.", "danger")
                return redirect(url_for("acts_list"))

            status = "draft" if is_draft else "completed"

            con.execute(
                """
                UPDATE acts SET
                    customer_name = ?,
                    customer_representative = ?,
                    device_model = ?,
                    serial_number = ?,
                    printeco_label = ?,
                    device_type = ?,
                    comment_label = ?,
                    address = ?,
                    phone = ?,
                    fault = ?,
                    counter_bw = ?,
                    counter_color = ?,
                    service_kind = ?,
                    diagnostics_result = ?,
                    works_text = ?,
                    materials_text = ?,
                    work_date = ?,
                    start_time = ?,
                    end_time = ?,
                    device_condition = ?,
                    customer_signatory = ?,
                    source_device_id = ?,
                    updated_at = ?,
                    status = ?,
                    number_type = ?
                WHERE id = ?
                """,
                (
                    clean(request.form.get("customer_name")),
                    clean(request.form.get("customer_representative")),
                    clean(request.form.get("device_model")),
                    clean(request.form.get("serial_number")),
                    clean(request.form.get("printeco_label")),
                    clean(request.form.get("device_type")),
                    clean(request.form.get("comment_label")),
                    clean(request.form.get("address")),
                    clean(request.form.get("phone")),
                    clean(request.form.get("fault")),
                    clean(request.form.get("counter_bw")),
                    clean(request.form.get("counter_color")),
                    clean(request.form.get("service_kind")),
                    clean(request.form.get("diagnostics_result")),
                    clean_multiline(request.form.get("works_text")),
                    clean_multiline(request.form.get("materials_text")),
                    clean(request.form.get("work_date")),
                    clean(request.form.get("start_time")),
                    clean(request.form.get("end_time")),
                    clean(request.form.get("device_condition")),
                    clean(request.form.get("customer_signatory")),
                    device_id,
                    now.isoformat(timespec="seconds"),
                    status,
                    number_type,
                    act_id,
                ),
            )

            flash("Акт обновлён.", "success")
            return redirect(url_for("act_view", act_id=act_id))

        else:
            status = "draft" if is_draft else "completed"

            act_number = None
            if not is_draft and status == "completed":
                act_number = next_act_number(
                    con, engineer["code"], now, status, number_type
                )

            cursor = con.execute(
                """
                INSERT INTO acts (
                    act_number, number_type, engineer_key, engineer_name,
                    engineer_code, created_at, updated_at, status,
                    customer_name, customer_representative, device_model,
                    serial_number, printeco_label, device_type, comment_label,
                    address, phone, fault, counter_bw, counter_color,
                    service_kind, diagnostics_result, works_text, materials_text,
                    work_date, start_time, end_time, device_condition,
                    customer_signatory, source_device_id
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    act_number,
                    number_type,
                    session["engineer_key"],
                    engineer["name"],
                    engineer["code"],
                    now.isoformat(timespec="seconds"),
                    now.isoformat(timespec="seconds"),
                    status,
                    clean(request.form.get("customer_name")),
                    clean(request.form.get("customer_representative")),
                    clean(request.form.get("device_model")),
                    clean(request.form.get("serial_number")),
                    clean(request.form.get("printeco_label")),
                    clean(request.form.get("device_type")),
                    clean(request.form.get("comment_label")),
                    clean(request.form.get("address")),
                    clean(request.form.get("phone")),
                    clean(request.form.get("fault")),
                    clean(request.form.get("counter_bw")),
                    clean(request.form.get("counter_color")),
                    clean(request.form.get("service_kind")),
                    clean(request.form.get("diagnostics_result")),
                    clean_multiline(request.form.get("works_text")),
                    clean_multiline(request.form.get("materials_text")),
                    clean(request.form.get("work_date")),
                    clean(request.form.get("start_time")),
                    clean(request.form.get("end_time")),
                    clean(request.form.get("device_condition")),
                    clean(request.form.get("customer_signatory")),
                    device_id,
                ),
            )

            flash("Акт сохранён.", "success")
            return redirect(url_for("act_view", act_id=cursor.lastrowid))

@app.route("/acts/<int:act_id>/delete", methods=["POST"])
def act_delete(act_id):
    if not require_engineer():
        return redirect(url_for("index"))

    with db() as con:
        con.execute("DELETE FROM acts WHERE id = ?", (act_id,))

    flash("Акт удалён.", "success")
    return redirect(url_for("acts_list"))

@app.route("/export")
def export_acts():
    if not require_engineer():
        return redirect(url_for("index"))

    q = request.args.get("q", "")
    engineer_filter = request.args.get("engineer", "")
    status_filter = request.args.get("status", "")
    date_sort = request.args.get("date_sort", "")
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")

    with db() as con:
        query = "SELECT * FROM acts WHERE 1=1"
        params = []

        if q:
            query += " AND (act_number LIKE ? OR serial_number LIKE ? OR customer_name LIKE ?)"
            params.extend([f"%{q}%"] * 3)

        if engineer_filter:
            query += " AND engineer_key = ?"
            params.append(engineer_filter)

        if status_filter:
            query += " AND status = ?"
            params.append(status_filter)

        if date_from:
            query += " AND DATE(created_at) >= ?"
            params.append(date_from)

        if date_to:
            query += " AND DATE(created_at) <= ?"
            params.append(date_to)

        if date_sort == "asc":
            query += " ORDER BY created_at ASC"
        elif date_sort == "desc":
            query += " ORDER BY created_at DESC"
        else:
            query += " ORDER BY created_at DESC"

        acts = con.execute(query, params).fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Акты"

    headers = [
        "Номер акта",
        "Статус",
        "Инженер",
        "Клиент",
        "Модель",
        "Серийный номер",
        "Дата создания",
        "Дата работы",
        "ЗИП",
    ]

    header_font = Font(bold=True)
    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")

    for col_num, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_num, act in enumerate(acts, start=2):
        ws.cell(row=row_num, column=1, value=act["act_number"] or "")
        ws.cell(row=row_num, column=2, value="Черновик" if act["status"] == "draft" else "Завершен")
        ws.cell(row=row_num, column=3, value=act["engineer_name"])
        ws.cell(row=row_num, column=4, value=act["customer_name"] or "")
        ws.cell(row=row_num, column=5, value=act["device_model"] or "")
        ws.cell(row=row_num, column=6, value=act["serial_number"] or "")
        ws.cell(row=row_num, column=7, value=act["created_at"][:16].replace("T", " "))
        ws.cell(row=row_num, column=8, value=act["work_date"] or "")
        ws.cell(row=row_num, column=9, value=act["materials_text"] or "")

    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 20

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"acts_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )

def parse_materials(text):
    if not text:
        return []

    items = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line:
            continue

        parts = line.split(";")
        item = {"name": "", "article": "", "quantity": ""}

        for part in parts:
            part = part.strip().lower()
            if part.startswith("наименование:"):
                item["name"] = part.replace("наименование:", "").strip()
            elif part.startswith("артикул:"):
                item["article"] = part.replace("артикул:", "").strip()
            elif part.startswith("кол-во:"):
                item["quantity"] = part.replace("кол-во:", "").strip()

        if item["name"] or item["article"] or item["quantity"]:
            items.append(item)

    return items

@app.route("/act/<int:act_id>/print")
def act_print(act_id):
    with db() as con:
        act = con.execute(
            "SELECT * FROM acts WHERE id = ?", (act_id,)
        ).fetchone()

    if not act:
        flash("Акт не найден.", "danger")
        return redirect(url_for("acts_list"))

    return render_template(
        "act_print.html",
        act=act,
        materials=parse_materials(act["materials_text"]),
    )

if __name__ == "__main__":
    init_db()
    try:
        import_google_sheet()
    except Exception as e:
        print(f"Warning: Could not import Google Sheet: {e}")
    app.run(debug=True, host="0.0.0.0", port=5000)
