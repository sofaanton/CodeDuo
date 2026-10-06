"""NaryadAI demo API: FastAPI + SQLite + WebSocket events.

Run from the project root with: uvicorn backend.main:app --host 0.0.0.0 --port 8000
"""
from __future__ import annotations

import asyncio
import csv
import io
import json
import random
import sqlite3
from contextlib import asynccontextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, Query, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from openpyxl import Workbook
from .ai_agent import evaluate_order

ROOT = Path(__file__).resolve().parent.parent
FRONTEND = ROOT / "frontend"
DB_PATH = Path(__file__).resolve().parent / "naryad.sqlite3"
NOW = lambda: datetime.now(timezone.utc)

SECTIONS = ["Участок дробления", "Обогатительная фабрика", "РМЦ", "Цех сушки"]
EQUIPMENT = [
    ("Дробилка КМД-1750", 0), ("Дробилка КСД-2200", 0), ("Конвейер К-3", 0),
    ("Конвейер К-7", 0), ("Грохот ГИЛ-52", 0), ("Питатель П-804", 0),
    ("Мельница ШМ-2", 1), ("Насос НС-25", 1), ("Сепаратор СБМ-1", 1),
    ("Флотомашина ФМ-12", 1), ("Конвейер К-12", 1), ("Сгуститель С-4", 1),
    ("Токарный станок 1К62", 2), ("Фрезерный станок 6Р12", 2),
    ("Кран-балка КБ-5", 2), ("Сварочный пост СП-2", 2), ("Компрессор ВК-7", 2),
    ("Экскаватор ЭКГ-10", 0), ("Экскаватор ЭКГ-8И", 0), ("Трансформатор ТМ-630", 1),
    ("Сушильный барабан СБ-3", 3), ("Вентилятор ВДН-15", 3),
    ("Циклон ЦН-15", 3), ("Нория Н-20", 3), ("Дымосос ДН-12", 3),
]
DEFECTS = [
    ("М-01", "Износ подшипника", "Механика"), ("М-02", "Несоосность вала", "Механика"),
    ("М-03", "Повреждение зубчатой передачи", "Механика"), ("М-04", "Ослабление крепежа", "Механика"),
    ("М-05", "Разрушение муфты", "Механика"), ("М-06", "Деформация рамы", "Механика"),
    ("М-07", "Износ футеровки", "Механика"), ("М-08", "Повреждение ленты", "Механика"),
    ("Э-01", "Короткое замыкание", "Электрика"), ("Э-02", "Перегрев двигателя", "Электрика"),
    ("Э-03", "Повреждение кабеля", "Электрика"), ("Э-04", "Отказ датчика", "Электрика"),
    ("Э-05", "Неисправность пускателя", "Электрика"), ("Э-06", "Потеря фазы", "Электрика"),
    ("Г-01", "Течь масла", "Гидравлика"), ("Г-02", "Износ уплотнения", "Гидравлика"),
    ("Г-03", "Отказ гидронасоса", "Гидравлика"), ("Г-04", "Засорение фильтра", "Гидравлика"),
    ("П-01", "Падение давления воздуха", "Пневматика"), ("П-02", "Утечка воздуха", "Пневматика"),
    ("С-01", "Недостаток смазки", "Смазка"), ("С-02", "Загрязнение смазки", "Смазка"),
    ("С-03", "Нарушение графика смазки", "Смазка"), ("С-04", "Отказ системы смазки", "Смазка"),
]
MATERIALS = [
    "Подшипник 3208", "Подшипник 6312", "Подшипник 22216", "Масло И-40", "Масло И-20А",
    "Смазка Литол-24", "Смазка ЦИАТИМ-221", "Сальник 50×70", "Сальник 80×100", "Манжета 40×60",
    "Рукав гидравлический 12 мм", "Рукав гидравлический 25 мм", "Фильтр масляный ФМ-12",
    "Фильтр воздушный ФВ-8", "Электрод УОНИ 13/55", "Электрод МР-3", "Кабель ВВГ 3×2.5",
    "Кабель КГ 4×6", "Контактор КМИ-22510", "Автомат ВА47-29", "Датчик вибрации ДВ-01",
    "Болт М16×60", "Гайка М16", "Шайба гровер М16", "Муфта МУВП-3", "Ремень клиновой А-1250",
    "Транспортерная лента 800 мм", "Скребок конвейерный", "Футеровка резиновая 10 мм",
    "Зубчатое колесо Z-42", "Кольцо уплотнительное 32 мм", "Краска грунтовочная",
    "Растворитель Р-4", "Проволока сварочная Св-08Г2С", "Круг отрезной 125 мм",
    "Шланг пневматический 10 мм", "Клапан обратный КО-25", "Паста монтажная",
    "Смазочный ниппель М10", "Вал приводной К-3", "Прокладка паронитовая 2 мм",
    "Лампа сигнальная 24В",
]
EXECUTORS = [
    ("Ахметов Ерлан", "Слесарь", 5), ("Смирнов Дмитрий", "Слесарь", 6),
    ("Петров Николай", "Слесарь", 4), ("Ким Виктор", "Электрик", 5),
    ("Иванова Алия", "Сварщик", 6), ("Сергеев Павел", "Слесарь", 5),
    ("Толеуов Марат", "Электрик", 4), ("Кузнецов Илья", "Слесарь", 6),
    ("Бекова Дана", "Сварщик", 5), ("Нургалиев Арман", "Слесарь", 4),
    ("Васильев Олег", "Электрик", 6), ("Жумабеков Али", "Слесарь", 5),
    ("Садыкова Лаура", "Сварщик", 4), ("Морозов Евгений", "Слесарь", 5),
    ("Омаров Руслан", "Электрик", 5),
]
PRIORITIES = ["Аварийный", "Высокий", "Обычный", "Плановый"]
STATUSES = ["Выдано", "Принят", "В работе", "В очереди", "На доработке", "Закрыт", "Просрочен"]


def connect() -> sqlite3.Connection:
    db = sqlite3.connect(DB_PATH, timeout=15)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA foreign_keys=ON")
    return db


def iso(dt: datetime | None = None) -> str:
    return (dt or NOW()).isoformat(timespec="seconds")


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with connect() as db:
        db.execute("PRAGMA journal_mode=WAL")
        db.executescript("""
        CREATE TABLE IF NOT EXISTS sections(id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL);
        CREATE TABLE IF NOT EXISTS equipment(id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, section_id INTEGER NOT NULL REFERENCES sections(id), active INTEGER NOT NULL DEFAULT 1);
        CREATE TABLE IF NOT EXISTS employees(id INTEGER PRIMARY KEY, name TEXT NOT NULL, role TEXT NOT NULL, specialty TEXT, grade INTEGER, active INTEGER NOT NULL DEFAULT 1);
        CREATE TABLE IF NOT EXISTS defect_codes(id INTEGER PRIMARY KEY, code TEXT UNIQUE NOT NULL, title TEXT NOT NULL, category TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS materials(id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, unit TEXT NOT NULL DEFAULT 'шт');
        CREATE TABLE IF NOT EXISTS orders(
          id INTEGER PRIMARY KEY AUTOINCREMENT, number INTEGER UNIQUE NOT NULL, job_type TEXT NOT NULL, description TEXT NOT NULL,
          section_id INTEGER NOT NULL REFERENCES sections(id), equipment_id INTEGER NOT NULL REFERENCES equipment(id),
          assignee_id INTEGER REFERENCES employees(id), priority TEXT NOT NULL, status TEXT NOT NULL,
          created_at TEXT NOT NULL, due_at TEXT NOT NULL, accepted_at TEXT, started_at TEXT, completed_at TEXT,
          reason TEXT, work_done TEXT, defect_code_id INTEGER REFERENCES defect_codes(id), materials_json TEXT NOT NULL DEFAULT '[]',
          before_photo TEXT, after_photo TEXT, comment TEXT, ai_score REAL, ai_verdict TEXT, ai_report TEXT,
          downtime_min INTEGER NOT NULL DEFAULT 0, rework_count INTEGER NOT NULL DEFAULT 0, warning_sent INTEGER NOT NULL DEFAULT 0, updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS order_events(id INTEGER PRIMARY KEY AUTOINCREMENT, order_id INTEGER NOT NULL REFERENCES orders(id), actor TEXT NOT NULL, action TEXT NOT NULL, note TEXT, created_at TEXT NOT NULL);
        CREATE INDEX IF NOT EXISTS idx_orders_status_due ON orders(status,due_at);
        CREATE INDEX IF NOT EXISTS idx_orders_equipment_date ON orders(equipment_id,created_at);
        CREATE INDEX IF NOT EXISTS idx_orders_assignee_date ON orders(assignee_id,created_at);
        """)
        if db.execute("SELECT count(*) FROM sections").fetchone()[0] > 0:
            return
        db.executemany("INSERT INTO sections(name) VALUES(?)", [(s,) for s in SECTIONS])
        section_ids = {r["name"]: r["id"] for r in db.execute("SELECT id,name FROM sections")}
        db.executemany("INSERT INTO equipment(name,section_id) VALUES(?,?)", [(name, section_ids[SECTIONS[index]]) for name, index in EQUIPMENT])
        employees = [("Елена Смирнова", "Мастер", None, None), ("Алексей Иванов", "Руководитель", None, None), ("Администратор системы", "Администратор", None, None)] + [(name, "Исполнитель", specialty, grade) for name, specialty, grade in EXECUTORS]
        db.executemany("INSERT INTO employees(name,role,specialty,grade) VALUES(?,?,?,?)", employees)
        db.executemany("INSERT INTO defect_codes(code,title,category) VALUES(?,?,?)", DEFECTS)
        db.executemany("INSERT INTO materials(name,unit) VALUES(?,?)", [(m, "л" if "Масло" in m or "Смазка" in m else "м" if "Кабель" in m or "лента" in m.lower() else "кг" if "Электрод" in m or "Проволока" in m else "шт") for m in MATERIALS])
        equip_rows = db.execute("SELECT id,name,section_id FROM equipment ORDER BY id").fetchall()
        exec_rows = db.execute("SELECT id,name,specialty FROM employees WHERE role='Исполнитель' ORDER BY id").fetchall()
        defect_rows = db.execute("SELECT id,code FROM defect_codes ORDER BY id").fetchall()
        material_rows = db.execute("SELECT id,name FROM materials ORDER BY id").fetchall()
        rng = random.Random(20261005)
        now = NOW()
        # Synthetic history with three seeded anomalies: K-3 repeats, Pетров has high rework, and post-PPR failures.
        rows = []
        for i in range(1, 501):
            age = rng.randint(0, 89)
            created = now - timedelta(days=age, hours=rng.randint(0, 20), minutes=rng.randint(0, 59))
            equip = rng.choice(equip_rows)
            if i <= 42:
                equip = next(e for e in equip_rows if e["name"] == "Конвейер К-3")
            worker = rng.choice(exec_rows)
            if 43 <= i <= 82:
                worker = next(e for e in exec_rows if e["name"] == "Петров Николай")
            if 83 <= i <= 102:
                equip = next(e for e in equip_rows if e["name"] == "Дробилка КМД-1750")
                created = now - timedelta(days=(i-83)//2, hours=rng.randint(0, 10))
            defect = next(d for d in defect_rows if d["code"] == "М-01") if i <= 42 else rng.choice(defect_rows)
            priority = rng.choices(PRIORITIES, weights=[5, 16, 55, 24])[0]
            job_type = "Внеплановый" if priority in ("Аварийный", "Высокий", "Обычный") else "Плановый"
            status = "Закрыт" if age > 0 or i > 488 else rng.choice(["В работе", "Выдано", "В очереди", "Просрочен"])
            rework = 1 if worker["name"] == "Петров Николай" and rng.random() < .42 else (1 if rng.random() < .06 else 0)
            due = created + timedelta(hours=rng.randint(1, 6))
            completed = created + timedelta(minutes=rng.randint(35, 240)) if status == "Закрыт" else None
            mats = [{"id": m["id"], "name": m["name"], "quantity": rng.randint(1, 3), "unit": "шт"} for m in rng.sample(material_rows, 1)] if status == "Закрыт" else []
            desc = "Повторный отказ после ППР: вибрация и нагрев редуктора" if 83 <= i <= 102 else "Повторная остановка, шум и нагрев узла" if equip["name"] == "Конвейер К-3" else "Плановый осмотр и устранение выявленного дефекта" if job_type == "Плановый" else "Вибрация, посторонний шум или течь масла; проверить узел и устранить дефект"
            rows.append((i, job_type, desc, equip["section_id"], equip["id"], worker["id"], priority, status, iso(created), iso(due), iso(created + timedelta(minutes=5)) if status != "Выдано" else None, iso(created + timedelta(minutes=12)) if status in ("Закрыт", "В работе", "Просрочен") else None, iso(completed) if completed else None, "Неисправность повторилась после ППР" if 83 <= i <= 102 else None, "Узел осмотрен, дефект устранён" if status == "Закрыт" else None, defect["id"], json.dumps(mats, ensure_ascii=False), None, None, "Демо-запись истории", round(rng.uniform(3.4, 5.0), 1) if status == "Закрыт" else None, "Принято" if status == "Закрыт" else None, "Синтетическая демонстрационная запись", rng.randint(20, 220) if status == "Закрыт" else 0, rework, iso(completed or created)))
        db.executemany("""INSERT INTO orders(number,job_type,description,section_id,equipment_id,assignee_id,priority,status,created_at,due_at,accepted_at,started_at,completed_at,reason,work_done,defect_code_id,materials_json,before_photo,after_photo,comment,ai_score,ai_verdict,ai_report,downtime_min,rework_count,updated_at)
          VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", rows)
        db.executemany("INSERT INTO order_events(order_id,actor,action,note,created_at) VALUES(?,?,?,?,?)", [(i, "Система", "Закрыт", "Начальная демонстрационная история", iso(now - timedelta(days=rng.randint(1, 89)))) for i in range(1, 501)])


class Hub:
    def __init__(self):
        self.clients: set[WebSocket] = set()

    async def connect(self, ws: WebSocket):
        await ws.accept()
        self.clients.add(ws)

    def disconnect(self, ws: WebSocket):
        self.clients.discard(ws)

    async def publish(self, payload: dict[str, Any]):
        stale = []
        for client in list(self.clients):
            try:
                await client.send_json(payload)
            except Exception:
                stale.append(client)
        for client in stale:
            self.disconnect(client)


hub = Hub()


def order_dict(row: sqlite3.Row) -> dict[str, Any]:
    item = dict(row)
    item["materials"] = json.loads(item.pop("materials_json") or "[]")
    item["job_type"] = item.pop("job_type")
    return item


def get_order(db: sqlite3.Connection, order_id: int):
    row = db.execute("""SELECT o.*,s.name section,e.name equipment,p.name assignee,p.specialty specialty,d.code defect_code,d.title defect_title
        FROM orders o JOIN sections s ON s.id=o.section_id JOIN equipment e ON e.id=o.equipment_id
        LEFT JOIN employees p ON p.id=o.assignee_id LEFT JOIN defect_codes d ON d.id=o.defect_code_id WHERE o.id=?""", (order_id,)).fetchone()
    return order_dict(row) if row else None


def event(db, order_id: int, actor: str, action: str, note: str | None = None):
    db.execute("INSERT INTO order_events(order_id,actor,action,note,created_at) VALUES(?,?,?,?,?)", (order_id, actor, action, note, iso()))


async def deadline_watch():
    while True:
        now = NOW(); warn_before = now + timedelta(minutes=30); broadcasts = []
        with connect() as db:
            candidates = db.execute("SELECT id,due_at,status,warning_sent FROM orders WHERE status NOT IN ('Закрыт','Отклонен')").fetchall()
            for row in candidates:
                try: due = datetime.fromisoformat(row["due_at"])
                except ValueError: continue
                if due.tzinfo is None: due = due.replace(tzinfo=timezone.utc)
                if due <= now and row["status"] != "Просрочен":
                    db.execute("UPDATE orders SET status='Просрочен',updated_at=? WHERE id=?", (iso(), row["id"]))
                    event(db, row["id"], "ИИ-контролёр", "Просрочен", "Автоматическая эскалация мастеру")
                    broadcasts.append({"type":"deadline.overdue","order_id":row["id"]})
                elif now < due <= warn_before and not row["warning_sent"]:
                    db.execute("UPDATE orders SET warning_sent=1,updated_at=? WHERE id=?", (iso(), row["id"]))
                    event(db, row["id"], "ИИ-контролёр", "Срок через 30 минут", "Напоминание исполнителю и мастеру")
                    broadcasts.append({"type":"deadline.warning","order_id":row["id"]})
            db.commit()
        for message in broadcasts: await hub.publish(message)
        await asyncio.sleep(30)


@asynccontextmanager
async def lifespan(_: FastAPI):
    init_db()
    task = asyncio.create_task(deadline_watch())
    try:
        yield
    finally:
        task.cancel()
        try: await task
        except asyncio.CancelledError: pass


app = FastAPI(title="НарядAI API", version="1.0.0", description="Демо API цифровых производственных нарядов", lifespan=lifespan)


class OrderCreate(BaseModel):
    description: str = Field(min_length=5, max_length=2000)
    job_type: str = "Внеплановый"
    section_id: int
    equipment_id: int
    assignee_id: int | None = None
    priority: str = "Обычный"
    due_at: str
    comment: str | None = None
    before_photo: str | None = None


class ActionBody(BaseModel):
    action: str
    actor: str = "Исполнитель"
    reason: str | None = None


class CompleteBody(BaseModel):
    actor: str = "Исполнитель"
    work_done: str = Field(min_length=3, max_length=3000)
    defect_code_id: int
    materials: list[dict[str, Any]] = Field(default_factory=list)
    after_photo: str | None = None
    comment: str | None = None


class CatalogCreate(BaseModel):
    name: str
    section_id: int | None = None
    specialty: str | None = None
    grade: int | None = None
    category: str | None = None
    unit: str | None = None
    code: str | None = None


class DemoLoginBody(BaseModel):
    email: str
    password: str


@app.get("/api/health")
def health():
    return {"status": "ok", "service": "НарядAI", "database": "SQLite"}


@app.post("/api/auth/demo-login")
def demo_login(body: DemoLoginBody):
    """Demo-only credential check for the hackathon presentation; not production auth."""
    accounts = {
        "master@km.kz": ("pass123", "Мастер", "Елена Смирнова"),
        "worker@km.kz": ("pass123", "Исполнитель", None),
        "admin@km.kz": ("pass123", "Руководитель", "Главный механик"),
    }
    email = body.email.strip().lower()
    account = accounts.get(email)
    if not account or body.password != account[0]:
        raise HTTPException(401, "Проверьте логин и пароль")
    password, role, display_name = account
    employee_id = None
    with connect() as db:
        if role == "Исполнитель":
            employee = db.execute("SELECT id,name FROM employees WHERE role='Исполнитель' AND active=1 ORDER BY id LIMIT 1").fetchone()
            if employee:
                employee_id, display_name = employee["id"], employee["name"]
        else:
            employee = db.execute("SELECT id FROM employees WHERE name=? ORDER BY id LIMIT 1", (display_name,)).fetchone()
            employee_id = employee["id"] if employee else None
    return {"email": email, "role": role, "name": display_name or role, "employee_id": employee_id}


@app.get("/api/catalog")
def catalog():
    with connect() as db:
        return {
            "sections": [dict(r) for r in db.execute("SELECT * FROM sections ORDER BY name")],
            "equipment": [dict(r) for r in db.execute("SELECT e.*,s.name section FROM equipment e JOIN sections s ON s.id=e.section_id WHERE e.active=1 ORDER BY e.name")],
            "employees": [dict(r) for r in db.execute("SELECT p.*,coalesce(x.open_count,0) open_count,coalesce(x.rating,4.5) rating FROM employees p LEFT JOIN (SELECT assignee_id,sum(CASE WHEN status NOT IN ('Закрыт','Отклонен') THEN 1 ELSE 0 END) open_count,avg(coalesce(ai_score,4.5)) rating FROM orders GROUP BY assignee_id) x ON x.assignee_id=p.id WHERE p.active=1 ORDER BY p.name")],
            "defects": [dict(r) for r in db.execute("SELECT * FROM defect_codes ORDER BY code")],
            "materials": [dict(r) for r in db.execute("SELECT * FROM materials ORDER BY name")],
            "norms": [{"job_type": "Замена подшипника", "hours": 2.0}, {"job_type": "Ремонт электродвигателя", "hours": 3.0}, {"job_type": "Устранение течи", "hours": 1.5}],
        }


@app.get("/api/orders")
def list_orders(status: str | None = None, section_id: int | None = None, assignee_id: int | None = None, priority: str | None = None, q: str | None = None, limit: int = Query(200, ge=1, le=1000)):
    sql = """SELECT o.*,s.name section,e.name equipment,p.name assignee,p.specialty specialty,d.code defect_code,d.title defect_title
        FROM orders o JOIN sections s ON s.id=o.section_id JOIN equipment e ON e.id=o.equipment_id
        LEFT JOIN employees p ON p.id=o.assignee_id LEFT JOIN defect_codes d ON d.id=o.defect_code_id WHERE 1=1"""
    args: list[Any] = []
    for column, value in (("o.status", status), ("o.section_id", section_id), ("o.assignee_id", assignee_id), ("o.priority", priority)):
        if value is not None:
            sql += f" AND {column}=?"; args.append(value)
    if q:
        sql += " AND (cast(o.number as text) LIKE ? OR e.name LIKE ? OR p.name LIKE ? OR o.description LIKE ?)"
        args.extend([f"%{q}%"] * 4)
    sql += " ORDER BY CASE o.priority WHEN 'Аварийный' THEN 0 WHEN 'Высокий' THEN 1 ELSE 2 END, o.due_at DESC LIMIT ?"
    args.append(limit)
    with connect() as db:
        return [order_dict(r) for r in db.execute(sql, args)]


@app.post("/api/orders", status_code=201)
async def create_order(body: OrderCreate):
    with connect() as db:
        if not db.execute("SELECT id FROM equipment WHERE id=? AND section_id=?", (body.equipment_id, body.section_id)).fetchone():
            raise HTTPException(422, "Оборудование не относится к выбранному участку")
        number = db.execute("SELECT coalesce(max(number),100)+1 FROM orders").fetchone()[0]
        cur = db.execute("INSERT INTO orders(number,job_type,description,section_id,equipment_id,assignee_id,priority,status,created_at,due_at,before_photo,comment,updated_at) VALUES(?,?,?,?,?,?,?,'Выдано',?,?,?,?,?)", (number, body.job_type, body.description, body.section_id, body.equipment_id, body.assignee_id, body.priority, iso(), body.due_at, body.before_photo, body.comment, iso()))
        order_id = cur.lastrowid
        event(db, order_id, "Мастер смены", "Выдан", body.comment)
        db.commit()
        result = get_order(db, order_id)
    await hub.publish({"type": "order.created", "order": result})
    return result


@app.post("/api/orders/{order_id}/action")
async def order_action(order_id: int, body: ActionBody):
    mapping = {"accept": "Принят", "queue": "В очереди", "reject": "Отклонен", "start": "В работе", "pause": "Приостановлен", "resume": "В работе"}
    if body.action not in mapping:
        raise HTTPException(400, "Неизвестное действие")
    if body.action in ("reject", "pause") and not (body.reason or "").strip():
        raise HTTPException(422, "Для отклонения или паузы укажите причину")
    with connect() as db:
        old = db.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
        if not old:
            raise HTTPException(404, "Наряд не найден")
        new_status = mapping[body.action]
        now = iso()
        db.execute("UPDATE orders SET status=?,reason=?,accepted_at=CASE WHEN ?='accept' THEN ? ELSE accepted_at END,started_at=CASE WHEN ?='start' OR ?='resume' THEN coalesce(started_at,?) ELSE started_at END,updated_at=? WHERE id=?", (new_status, body.reason, body.action, now, body.action, body.action, now, now, order_id))
        event(db, order_id, body.actor, body.action, body.reason)
        db.commit(); result = get_order(db, order_id)
    await hub.publish({"type": "order.updated", "order": result})
    return result


@app.post("/api/orders/{order_id}/complete")
async def complete_order(order_id: int, body: CompleteBody):
    with connect() as db:
        if not db.execute("SELECT id FROM orders WHERE id=?", (order_id,)).fetchone():
            raise HTTPException(404, "Наряд не найден")
        db.execute("UPDATE orders SET work_done=?,defect_code_id=?,materials_json=?,after_photo=?,comment=?,completed_at=?,status='На доработке',updated_at=? WHERE id=?", (body.work_done, body.defect_code_id, json.dumps(body.materials, ensure_ascii=False), body.after_photo, body.comment, iso(), iso(), order_id))
        order = get_order(db, order_id)
        report = evaluate_order(order)
        score, issues, verdict = report["score"], report["issues"], report["verdict"]
        db.execute("UPDATE orders SET ai_score=?,ai_verdict=?,ai_report=?,status='На доработке',updated_at=? WHERE id=?", (score, verdict, json.dumps(report, ensure_ascii=False), iso(), order_id))
        event(db, order_id, body.actor, "Отчёт отправлен на проверку", verdict)
        db.commit(); result = get_order(db, order_id)
    await hub.publish({"type": "order.completed", "order": result, "ai": report})
    return {"order": result, "ai": report}


@app.post("/api/orders/{order_id}/review")
async def review_order(order_id: int, accept: bool, actor: str = "Мастер смены", score: float | None = Query(None, ge=1, le=5), note: str | None = None):
    with connect() as db:
        if not db.execute("SELECT id FROM orders WHERE id=?", (order_id,)).fetchone():
            raise HTTPException(404, "Наряд не найден")
        status = "Закрыт" if accept else "В работе"
        db.execute("UPDATE orders SET status=?,ai_score=coalesce(?,ai_score),reason=?,rework_count=rework_count+?,updated_at=? WHERE id=?", (status, score, None if accept else (note or "Требуется доработка"), 0 if accept else 1, iso(), order_id))
        event(db, order_id, actor, "Принят мастером" if accept else "Возвращён на доработку", note)
        db.commit(); result = get_order(db, order_id)
    await hub.publish({"type": "order.reviewed", "order": result})
    return result


@app.get("/api/ai/assign")
def suggest_assignee(equipment_id: int, specialty: str | None = None):
    with connect() as db:
        equipment = db.execute("SELECT name FROM equipment WHERE id=?", (equipment_id,)).fetchone()
        if not equipment: raise HTTPException(404, "Оборудование не найдено")
        workers = db.execute("""SELECT p.id,p.name,p.specialty,p.grade,coalesce(a.open_count,0) open_count,coalesce(r.rating,4.5) rating
          FROM employees p LEFT JOIN (SELECT assignee_id,count(*) open_count FROM orders WHERE status IN ('Выдано','Принят','В работе','В очереди') GROUP BY assignee_id) a ON a.assignee_id=p.id
          LEFT JOIN (SELECT assignee_id,avg(coalesce(ai_score,4.5)) rating FROM orders WHERE status='Закрыт' GROUP BY assignee_id) r ON r.assignee_id=p.id
          WHERE p.role='Исполнитель' AND p.active=1 ORDER BY p.id""").fetchall()
    if not specialty:
        equipment_name = equipment["name"].lower()
        specialty = "Электрик" if any(x in equipment_name for x in ("трансформатор", "электро", "кабель")) else "Сварщик" if "свар" in equipment_name else "Слесарь"
    def score(w):
        skill = 1 if specialty and w["specialty"] == specialty else 0
        load = 1 / (1 + w["open_count"])
        return .45 * (w["rating"] / 5) + .35 * load + .2 * skill
    candidates = sorted([dict(w) | {"recommendation_score": round(score(w), 3), "available": w["open_count"] == 0} for w in workers], key=lambda w: (w["open_count"] > 0, -w["recommendation_score"]))
    return {"equipment": equipment["name"], "suggestions": candidates[:5], "explanation": "Рекомендация учитывает профиль, текущую загрузку и среднюю оценку закрытых работ."}


@app.get("/api/ai/report/{order_id}")
def ai_report(order_id: int):
    with connect() as db:
        row = db.execute("SELECT ai_report,ai_score,ai_verdict,work_done,description,downtime_min FROM orders WHERE id=?", (order_id,)).fetchone()
        if not row: raise HTTPException(404, "Наряд не найден")
        details = json.loads(row["ai_report"] or "{}")
        return {"order_id": order_id, "score": row["ai_score"], "verdict": row["ai_verdict"], "details": details, "employee_summary": row["work_done"] or "Отчёт ещё не отправлен", "downtime_minutes": row["downtime_min"]}


@app.get("/api/orders/{order_id}/events")
def order_events(order_id: int):
    with connect() as db:
        if not db.execute("SELECT id FROM orders WHERE id=?", (order_id,)).fetchone():
            raise HTTPException(404, "Наряд не найден")
        return [dict(r) for r in db.execute("SELECT actor,action,note,created_at FROM order_events WHERE order_id=? ORDER BY id", (order_id,))]


@app.get("/api/dashboard")
def dashboard():
    with connect() as db:
        stats = {r["status"]: r["n"] for r in db.execute("SELECT status,count(*) n FROM orders WHERE created_at>=? GROUP BY status", (iso(NOW()-timedelta(days=1)),))}
        workforce = [dict(r) for r in db.execute("""SELECT p.id,p.name,p.specialty,p.grade,coalesce(a.open_count,0) open_count,coalesce(a.status,'Свободен') status,coalesce(r.rating,4.5) rating
          FROM employees p LEFT JOIN (SELECT assignee_id,count(*) open_count,max(status) status FROM orders WHERE status IN ('Выдано','Принят','В работе','В очереди','На доработке') GROUP BY assignee_id) a ON a.assignee_id=p.id
          LEFT JOIN (SELECT assignee_id,avg(coalesce(ai_score,4.5)) rating FROM orders WHERE status='Закрыт' GROUP BY assignee_id) r ON r.assignee_id=p.id WHERE p.role='Исполнитель' AND p.active=1 ORDER BY p.name""")]
        risk = [order_dict(r) for r in db.execute("""SELECT o.*,s.name section,e.name equipment,p.name assignee,p.specialty specialty,d.code defect_code,d.title defect_title FROM orders o JOIN sections s ON s.id=o.section_id JOIN equipment e ON e.id=o.equipment_id LEFT JOIN employees p ON p.id=o.assignee_id LEFT JOIN defect_codes d ON d.id=o.defect_code_id WHERE o.status NOT IN ('Закрыт','Отклонен') AND o.due_at<? ORDER BY o.due_at LIMIT 10""", (iso(),))]
        return {"shift": {"issued": sum(stats.values()), "working": stats.get("В работе", 0), "completed": stats.get("Закрыт", 0), "overdue": len(risk)}, "workforce": workforce, "risks": risk, "active_orders": [order_dict(r) for r in db.execute("""SELECT o.*,s.name section,e.name equipment,p.name assignee,p.specialty specialty,d.code defect_code,d.title defect_title FROM orders o JOIN sections s ON s.id=o.section_id JOIN equipment e ON e.id=o.equipment_id LEFT JOIN employees p ON p.id=o.assignee_id LEFT JOIN defect_codes d ON d.id=o.defect_code_id WHERE o.status NOT IN ('Закрыт','Отклонен') ORDER BY CASE o.priority WHEN 'Аварийный' THEN 0 WHEN 'Высокий' THEN 1 ELSE 2 END,o.due_at LIMIT 30""") ]}


@app.get("/api/analytics")
def analytics(days: int = Query(30, ge=1, le=365)):
    cutoff = iso(NOW() - timedelta(days=days))
    with connect() as db:
        summary = dict(db.execute("SELECT count(*) total,sum(CASE WHEN job_type='Внеплановый' THEN 1 ELSE 0 END) unplanned,sum(downtime_min) downtime,avg(ai_score) avg_score,sum(rework_count) reworks FROM orders WHERE created_at>=?", (cutoff,)).fetchone())
        by_equipment = [dict(r) for r in db.execute("SELECT e.name equipment,count(*) failures,sum(o.downtime_min) downtime FROM orders o JOIN equipment e ON e.id=o.equipment_id WHERE o.created_at>=? AND o.job_type='Внеплановый' GROUP BY e.id ORDER BY failures DESC LIMIT 8", (cutoff,))]
        by_day = [dict(r) for r in db.execute("SELECT substr(created_at,1,10) day,sum(CASE WHEN job_type='Внеплановый' THEN 1 ELSE 0 END) failures,sum(downtime_min) downtime FROM orders WHERE created_at>=? GROUP BY day ORDER BY day", (cutoff,))]
        by_worker = [dict(r) for r in db.execute("SELECT p.name employee,count(*) closed,avg(o.ai_score) rating,sum(o.rework_count) reworks,avg(CASE WHEN o.completed_at<=o.due_at THEN 100.0 ELSE 0.0 END) on_time FROM orders o JOIN employees p ON p.id=o.assignee_id WHERE o.created_at>=? AND o.status='Закрыт' GROUP BY p.id ORDER BY rating DESC", (cutoff,))]
        repeat_codes = [dict(r) for r in db.execute("SELECT e.name equipment,d.code defect_code,d.title defect,count(*) repeats FROM orders o JOIN equipment e ON e.id=o.equipment_id JOIN defect_codes d ON d.id=o.defect_code_id WHERE o.created_at>=? AND o.job_type='Внеплановый' GROUP BY o.equipment_id,o.defect_code_id HAVING repeats>=3 ORDER BY repeats DESC LIMIT 8", (cutoff,))]
        post_ppr = [dict(r) for r in db.execute("SELECT e.name equipment,count(*) failures FROM orders o JOIN equipment e ON e.id=o.equipment_id WHERE o.created_at>=? AND o.job_type='Внеплановый' AND o.description LIKE '%после ППР%' GROUP BY o.equipment_id ORDER BY failures DESC", (cutoff,))]
    for worker in by_worker:
        quality = (worker["rating"] or 0) / 5 * 100
        on_time = worker["on_time"] or 0
        volume = min(100, worker["closed"] * 4)
        return_penalty = min(100, (worker["reworks"] or 0) * 10)
        worker["score"] = round(quality * .4 + on_time * .3 + volume * .2 - return_penalty * .1, 1)
    anomalies = []
    if by_equipment and by_equipment[0]["failures"] >= 3:
        anomalies.append({"severity": "high", "title": f"Частые отказы: {by_equipment[0]['equipment']}", "detail": f"{by_equipment[0]['failures']} внеплановых остановок за {days} дн.; простой {by_equipment[0]['downtime']} мин.", "recommendation": "Проверьте узел и рассмотрите внеочередное обслуживание."})
    if repeat_codes:
        x = repeat_codes[0]
        anomalies.append({"severity": "high", "title": f"Повторный дефект {x['defect_code']} · {x['equipment']}", "detail": f"Повторился {x['repeats']} раз(а): {x['defect']}.", "recommendation": "Проверьте первопричину ремонта и соосность сопряжённых узлов."})
    if by_worker:
        weak = min(by_worker, key=lambda r: (r["rating"] or 0) - (r["reworks"] or 0) * .4)
        if weak["reworks"]:
            anomalies.append({"severity": "medium", "title": f"Возвраты на доработку · {weak['employee']}", "detail": f"{weak['reworks']} возвратов среди {weak['closed']} закрытых нарядов.", "recommendation": "Провести выборочную проверку закрытий и уточнить чек-лист качества."})
    return {"period_days": days, "summary": summary, "equipment": by_equipment, "timeline": by_day, "employees": by_worker, "repeat_codes": repeat_codes, "post_ppr": post_ppr, "anomalies": anomalies}


@app.post("/api/ai/chat")
def assistant(payload: dict[str, Any]):
    query = str(payload.get("query", "")).lower()
    with connect() as db:
        if any(k in query for k in ("свобод", "available", "бос")):
            workers = [r["name"] for r in db.execute("""SELECT p.name FROM employees p LEFT JOIN orders o ON o.assignee_id=p.id AND o.status IN ('Выдано','Принят','В работе','В очереди') WHERE p.role='Исполнитель' GROUP BY p.id HAVING count(o.id)=0 ORDER BY p.name LIMIT 8""")]
            reply = "Свободны: " + (", ".join(workers) if workers else "свободных исполнителей сейчас нет")
        elif any(k in query for k in ("просроч", "overdue", "кешігу")):
            count = db.execute("SELECT count(*) FROM orders WHERE status NOT IN ('Закрыт','Отклонен') AND due_at<?", (iso(),)).fetchone()[0]
            reply = f"Сейчас нарядов с истёкшим сроком: {count}. Откройте список рисков для переназначения."
        else:
            reply = "Могу найти свободных исполнителей или собрать сводку по просрочкам. Спросите: «кто свободен?» или «что просрочено?»"
    return {"reply": reply}


@app.get("/api/exports/orders.csv")
def export_orders(days: int = Query(30, ge=1, le=365)):
    with connect() as db:
        rows = db.execute("""SELECT o.number,o.job_type,o.description,s.name section,e.name equipment,p.name assignee,o.priority,o.status,o.created_at,o.due_at,o.completed_at,d.code defect_code,o.ai_score,o.downtime_min FROM orders o JOIN sections s ON s.id=o.section_id JOIN equipment e ON e.id=o.equipment_id LEFT JOIN employees p ON p.id=o.assignee_id LEFT JOIN defect_codes d ON d.id=o.defect_code_id WHERE o.created_at>=? ORDER BY o.created_at DESC""", (iso(NOW()-timedelta(days=days)),)).fetchall()
    output = io.StringIO(); writer = csv.writer(output, delimiter=";")
    writer.writerow(rows[0].keys() if rows else ["Наряды"])
    writer.writerows([tuple(r) for r in rows])
    return StreamingResponse(iter(["\ufeff" + output.getvalue()]), media_type="text/csv; charset=utf-8", headers={"Content-Disposition": f"attachment; filename=naryad-report-{days}d.csv"})


@app.get("/api/exports/orders.xlsx")
def export_orders_xlsx(days: int = Query(30, ge=1, le=365)):
    with connect() as db:
        rows = db.execute("""SELECT o.number,o.job_type,o.description,s.name section,e.name equipment,p.name assignee,o.priority,o.status,o.created_at,o.due_at,o.completed_at,d.code defect_code,o.ai_score,o.downtime_min FROM orders o JOIN sections s ON s.id=o.section_id JOIN equipment e ON e.id=o.equipment_id LEFT JOIN employees p ON p.id=o.assignee_id LEFT JOIN defect_codes d ON d.id=o.defect_code_id WHERE o.created_at>=? ORDER BY o.created_at DESC""", (iso(NOW()-timedelta(days=days)),)).fetchall()
    workbook = Workbook(); sheet = workbook.active; sheet.title = "Наряды"
    sheet.append(list(rows[0].keys()) if rows else ["Наряды"])
    for row in rows: sheet.append(list(row))
    sheet.freeze_panes = "A2"; sheet.auto_filter.ref = sheet.dimensions
    for column in sheet.columns:
        letter = column[0].column_letter
        sheet.column_dimensions[letter].width = min(42, max(13, max(len(str(cell.value or "")) for cell in column) + 2))
    buf = io.BytesIO(); workbook.save(buf); buf.seek(0)
    return StreamingResponse(buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=naryad-report-{days}d.xlsx"})


@app.post("/api/admin/catalog/{kind}")
def add_catalog(kind: str, item: CatalogCreate):
    tables = {"sections": ("INSERT INTO sections(name) VALUES(?)", (item.name,)), "equipment": ("INSERT INTO equipment(name,section_id) VALUES(?,?)", (item.name, item.section_id)), "employees": ("INSERT INTO employees(name,role,specialty,grade) VALUES(?,?,?,?)", (item.name, "Исполнитель", item.specialty, item.grade)), "materials": ("INSERT INTO materials(name,unit) VALUES(?,?)", (item.name, item.unit or "шт")), "defects": ("INSERT INTO defect_codes(code,title,category) VALUES(?,?,?)", (item.code or item.name, item.name, item.category or "Механика"))}
    if kind not in tables: raise HTTPException(404, "Неизвестный справочник")
    query, values = tables[kind]
    try:
        with connect() as db:
            cur = db.execute(query, values); db.commit(); new_id = cur.lastrowid
    except sqlite3.IntegrityError as exc:
        raise HTTPException(409, f"Не удалось добавить запись: {exc}")
    return {"id": new_id, "name": item.name}


@app.websocket("/ws")
async def websocket_endpoint(ws: WebSocket):
    await hub.connect(ws)
    try:
        await ws.send_json({"type": "connected", "time": iso(), "clients": len(hub.clients)})
        while True:
            await ws.receive_text()
    except WebSocketDisconnect:
        hub.disconnect(ws)


@app.get("/")
def root():
    return FileResponse(FRONTEND / "index.html")


@app.get("/manifest.webmanifest")
def manifest():
    return FileResponse(ROOT / "manifest.webmanifest", media_type="application/manifest+json")


@app.get("/sw.js")
def service_worker():
    return FileResponse(ROOT / "sw.js", media_type="application/javascript", headers={"Service-Worker-Allowed": "/"})


app.mount("/frontend", StaticFiles(directory=FRONTEND), name="frontend")
