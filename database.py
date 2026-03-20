"""
Database module for Glorious Pearls Complex School Quiz System.
Uses PostgreSQL on Railway, falls back to SQLite locally.
"""
import os, json, datetime

# Detect environment
DATABASE_URL = os.environ.get("DATABASE_URL", "")

def get_db():
    """Get database connection - PostgreSQL on Railway, SQLite locally"""
    if DATABASE_URL and DATABASE_URL.startswith("postgresql"):
        try:
            import psycopg2
            conn = psycopg2.connect(DATABASE_URL, sslmode="require")
            return conn, "postgres"
        except Exception as e:
            print("PostgreSQL error:", e)
    # Fall back to SQLite
    import sqlite3
    conn = sqlite3.connect("quiz_data.db")
    conn.row_factory = sqlite3.Row
    return conn, "sqlite"

def init_db():
    """Create all tables if they don't exist"""
    conn, db_type = get_db()
    cur = conn.cursor()

    if db_type == "postgres":
        # PostgreSQL syntax
        cur.execute("""
            CREATE TABLE IF NOT EXISTS results (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL,
                class TEXT,
                subject TEXT,
                assessment_type TEXT,
                assessment_label TEXT,
                strand TEXT,
                sub_strand TEXT,
                score INTEGER,
                total INTEGER,
                percentage REAL,
                grade TEXT,
                remark TEXT,
                date TEXT,
                time TEXT,
                details JSONB,
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS questions (
                id SERIAL PRIMARY KEY,
                subject TEXT NOT NULL,
                data JSONB NOT NULL,
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS homework_saves (
                id SERIAL PRIMARY KEY,
                save_key TEXT UNIQUE NOT NULL,
                data JSONB NOT NULL,
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS bece_questions (
                id SERIAL PRIMARY KEY,
                subject TEXT NOT NULL,
                year TEXT NOT NULL,
                data JSONB NOT NULL,
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS assignments (
                id SERIAL PRIMARY KEY,
                class TEXT NOT NULL,
                title TEXT NOT NULL,
                subject TEXT NOT NULL,
                assessment_type TEXT NOT NULL,
                due_date TEXT,
                assigned_date TEXT,
                assigned_by TEXT,
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS question_bank (
                id SERIAL PRIMARY KEY,
                subject TEXT NOT NULL,
                data JSONB NOT NULL,
                created_at TIMESTAMP DEFAULT NOW()
            )
        """)
    else:
        # SQLite syntax
        cur.execute("""
            CREATE TABLE IF NOT EXISTS results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT, class TEXT, subject TEXT,
                assessment_type TEXT, assessment_label TEXT,
                strand TEXT, sub_strand TEXT,
                score INTEGER, total INTEGER, percentage REAL,
                grade TEXT, remark TEXT, date TEXT, time TEXT,
                details TEXT, created_at TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS questions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                subject TEXT NOT NULL,
                data TEXT NOT NULL,
                created_at TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS homework_saves (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                save_key TEXT UNIQUE NOT NULL,
                data TEXT NOT NULL,
                created_at TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS bece_questions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                subject TEXT NOT NULL,
                year TEXT NOT NULL,
                data TEXT NOT NULL,
                created_at TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS assignments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                class TEXT NOT NULL,
                title TEXT NOT NULL,
                subject TEXT NOT NULL,
                assessment_type TEXT NOT NULL,
                due_date TEXT,
                assigned_date TEXT,
                assigned_by TEXT,
                created_at TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS question_bank (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                subject TEXT NOT NULL,
                data TEXT NOT NULL,
                created_at TEXT
            )
        """)

    conn.commit()
    cur.close()
    conn.close()
    print("Database initialised ({})".format(db_type))

# ── Results ────────────────────────────────────────────────────────

def db_load_results():
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM results ORDER BY id")
        rows = cur.fetchall()
        cur.close(); conn.close()
        results = []
        for row in rows:
            if db_type == "postgres":
                r = dict(row)
                r["details"] = r.get("details") or []
            else:
                r = dict(row)
                r["details"] = json.loads(r.get("details") or "[]")
            results.append(r)
        return results
    except Exception as e:
        print("db_load_results error:", e)
        return _file_load_results()

def db_save_result(result):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        details = result.get("details", [])
        details_str = json.dumps(details) if db_type == "sqlite" else json.dumps(details)
        if db_type == "postgres":
            import psycopg2.extras
            cur.execute("""
                INSERT INTO results (name, class, subject, assessment_type,
                    assessment_label, strand, sub_strand, score, total,
                    percentage, grade, remark, date, time, details)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (result.get("name",""), result.get("class",""),
                  result.get("subject",""), result.get("assessment_type",""),
                  result.get("assessment_label",""), result.get("strand",""),
                  result.get("sub_strand",""), result.get("score",0),
                  result.get("total",0), result.get("percentage",0),
                  result.get("grade","F"), result.get("remark",""),
                  result.get("date",""), result.get("time",""),
                  psycopg2.extras.Json(details)))
        else:
            cur.execute("""
                INSERT INTO results (name, class, subject, assessment_type,
                    assessment_label, strand, sub_strand, score, total,
                    percentage, grade, remark, date, time, details, created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (result.get("name",""), result.get("class",""),
                  result.get("subject",""), result.get("assessment_type",""),
                  result.get("assessment_label",""), result.get("strand",""),
                  result.get("sub_strand",""), result.get("score",0),
                  result.get("total",0), result.get("percentage",0),
                  result.get("grade","F"), result.get("remark",""),
                  result.get("date",""), result.get("time",""),
                  details_str,
                  datetime.datetime.now().isoformat()))
        conn.commit(); cur.close(); conn.close()
        # Also save backup to file
        _file_save_result(result)
    except Exception as e:
        print("db_save_result error:", e)
        _file_save_result(result)

def db_delete_result(idx):
    try:
        results = db_load_results()
        if 0 <= idx < len(results):
            row_id = results[idx].get("id")
            if row_id:
                conn, db_type = get_db()
                cur = conn.cursor()
                if db_type == "postgres":
                    cur.execute("DELETE FROM results WHERE id=%s", (row_id,))
                else:
                    cur.execute("DELETE FROM results WHERE id=?", (row_id,))
                conn.commit(); cur.close(); conn.close()
    except Exception as e:
        print("db_delete_result error:", e)

# ── Questions ──────────────────────────────────────────────────────

def db_load_questions(subject):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if db_type == "postgres":
            cur.execute("SELECT data FROM questions WHERE subject=%s ORDER BY id", (subject,))
        else:
            cur.execute("SELECT data FROM questions WHERE subject=? ORDER BY id", (subject,))
        rows = cur.fetchall(); cur.close(); conn.close()
        if not rows:
            return _file_load_questions(subject)
        qs = []
        for row in rows:
            d = row[0] if db_type == "postgres" else json.loads(row[0])
            if isinstance(d, list):
                qs.extend(d)
            else:
                qs.append(d)
        return qs
    except Exception as e:
        print("db_load_questions error:", e)
        return _file_load_questions(subject)

def db_save_questions(subject, questions):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        # Delete existing and reinsert all
        if db_type == "postgres":
            cur.execute("DELETE FROM questions WHERE subject=%s", (subject,))
            import psycopg2.extras
            cur.execute("INSERT INTO questions (subject, data) VALUES (%s,%s)",
                       (subject, psycopg2.extras.Json(questions)))
        else:
            cur.execute("DELETE FROM questions WHERE subject=?", (subject,))
            cur.execute("INSERT INTO questions (subject, data, created_at) VALUES (?,?,?)",
                       (subject, json.dumps(questions), datetime.datetime.now().isoformat()))
        conn.commit(); cur.close(); conn.close()
        _file_save_questions(subject, questions)
    except Exception as e:
        print("db_save_questions error:", e)
        _file_save_questions(subject, questions)

# ── Homework saves ─────────────────────────────────────────────────

def db_load_hw():
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        cur.execute("SELECT save_key, data FROM homework_saves")
        rows = cur.fetchall(); cur.close(); conn.close()
        saves = {}
        for row in rows:
            key = row[0]
            data = row[1] if db_type == "postgres" else json.loads(row[1])
            saves[key] = data
        return saves
    except Exception as e:
        print("db_load_hw error:", e)
        return _file_load_hw()

def db_save_hw(key, data):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if db_type == "postgres":
            import psycopg2.extras
            cur.execute("""
                INSERT INTO homework_saves (save_key, data)
                VALUES (%s,%s)
                ON CONFLICT (save_key) DO UPDATE SET data=EXCLUDED.data
            """, (key, psycopg2.extras.Json(data)))
        else:
            cur.execute("""
                INSERT OR REPLACE INTO homework_saves (save_key, data, created_at)
                VALUES (?,?,?)
            """, (key, json.dumps(data), datetime.datetime.now().isoformat()))
        conn.commit(); cur.close(); conn.close()
    except Exception as e:
        print("db_save_hw error:", e)
        _file_save_hw(key, data)

def db_del_hw(key):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if db_type == "postgres":
            cur.execute("DELETE FROM homework_saves WHERE save_key=%s", (key,))
        else:
            cur.execute("DELETE FROM homework_saves WHERE save_key=?", (key,))
        conn.commit(); cur.close(); conn.close()
    except Exception as e:
        print("db_del_hw error:", e)

# ── BECE questions ─────────────────────────────────────────────────

def db_load_bece(subject, year):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if db_type == "postgres":
            cur.execute("SELECT data FROM bece_questions WHERE subject=%s AND year=%s", (subject,year))
        else:
            cur.execute("SELECT data FROM bece_questions WHERE subject=? AND year=?", (subject,year))
        row = cur.fetchone(); cur.close(); conn.close()
        if not row:
            return _file_load_bece(subject, year)
        return row[0] if db_type=="postgres" else json.loads(row[0])
    except Exception as e:
        print("db_load_bece error:", e)
        return _file_load_bece(subject, year)

def db_save_bece(subject, year, questions):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if db_type == "postgres":
            import psycopg2.extras
            cur.execute("""
                INSERT INTO bece_questions (subject, year, data)
                VALUES (%s,%s,%s)
                ON CONFLICT DO NOTHING
            """, (subject, year, psycopg2.extras.Json(questions)))
            cur.execute("""
                UPDATE bece_questions SET data=%s WHERE subject=%s AND year=%s
            """, (psycopg2.extras.Json(questions), subject, year))
        else:
            cur.execute("DELETE FROM bece_questions WHERE subject=? AND year=?", (subject,year))
            cur.execute("INSERT INTO bece_questions (subject, year, data, created_at) VALUES (?,?,?,?)",
                       (subject, year, json.dumps(questions), datetime.datetime.now().isoformat()))
        conn.commit(); cur.close(); conn.close()
        _file_save_bece(subject, year, questions)
    except Exception as e:
        print("db_save_bece error:", e)
        _file_save_bece(subject, year, questions)

# ── File fallbacks ─────────────────────────────────────────────────

import threading
_lock = threading.Lock()
RESULTS_FILE  = "results.json"
HOMEWORK_FILE = "homework_saves.json"

def _file_load_results():
    return json.load(open(RESULTS_FILE)) if os.path.exists(RESULTS_FILE) else []

def _file_save_result(r):
    with _lock:
        rs = _file_load_results(); rs.append(r)
        json.dump(rs, open(RESULTS_FILE,"w"), indent=2)

def _file_load_questions(subject):
    f = "questions_{}.json".format(subject.replace(" ","_").replace("/","_"))
    return json.load(open(f)) if os.path.exists(f) else []

def _file_save_questions(subject, qs):
    f = "questions_{}.json".format(subject.replace(" ","_").replace("/","_"))
    json.dump(qs, open(f,"w"), indent=2)

def _file_load_hw():
    return json.load(open(HOMEWORK_FILE)) if os.path.exists(HOMEWORK_FILE) else {}

def _file_save_hw(key, data):
    saves = _file_load_hw(); saves[key] = data
    json.dump(saves, open(HOMEWORK_FILE,"w"), indent=2)

def _file_load_bece(subject, year):
    f = "bece_{}_{}.json".format(subject.replace(" ","_"), year)
    return json.load(open(f)) if os.path.exists(f) else []

def _file_save_bece(subject, year, qs):
    f = "bece_{}_{}.json".format(subject.replace(" ","_"), year)
    json.dump(qs, open(f,"w"), indent=2)

# ── Assignments ────────────────────────────────────────────────────

def db_load_assignments(class_name):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if db_type == "postgres":
            cur.execute("SELECT * FROM assignments WHERE class=%s ORDER BY due_date ASC", (class_name,))
        else:
            cur.execute("SELECT * FROM assignments WHERE class=? ORDER BY due_date ASC", (class_name,))
        rows = cur.fetchall(); cur.close(); conn.close()
        return [dict(r) for r in rows]
    except Exception as e:
        print("db_load_assignments error:", e)
        return []

def db_save_assignment(a):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if db_type == "postgres":
            cur.execute("""
                INSERT INTO assignments (class, title, subject, assessment_type, due_date, assigned_date, assigned_by)
                VALUES (%s,%s,%s,%s,%s,%s,%s)
            """, (a["class"], a["title"], a["subject"], a["assessment_type"],
                  a.get("due_date",""), a.get("assigned_date",""), a.get("assigned_by","")))
        else:
            cur.execute("""
                INSERT INTO assignments (class, title, subject, assessment_type, due_date, assigned_date, assigned_by, created_at)
                VALUES (?,?,?,?,?,?,?,?)
            """, (a["class"], a["title"], a["subject"], a["assessment_type"],
                  a.get("due_date",""), a.get("assigned_date",""), a.get("assigned_by",""),
                  datetime.datetime.now().isoformat()))
        conn.commit(); cur.close(); conn.close()
    except Exception as e:
        print("db_save_assignment error:", e)

def db_delete_assignment(aid):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if db_type == "postgres":
            cur.execute("DELETE FROM assignments WHERE id=%s", (aid,))
        else:
            cur.execute("DELETE FROM assignments WHERE id=?", (aid,))
        conn.commit(); cur.close(); conn.close()
    except Exception as e:
        print("db_delete_assignment error:", e)

# ── Question Bank ──────────────────────────────────────────────────

def db_load_bank(subject=None):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if subject:
            if db_type == "postgres":
                cur.execute("SELECT id, subject, data FROM question_bank WHERE subject=%s ORDER BY id", (subject,))
            else:
                cur.execute("SELECT id, subject, data FROM question_bank WHERE subject=? ORDER BY id", (subject,))
        else:
            cur.execute("SELECT id, subject, data FROM question_bank ORDER BY subject, id")
        rows = cur.fetchall(); cur.close(); conn.close()
        result = []
        for row in rows:
            if db_type == "postgres":
                d = dict(row["data"]) if isinstance(row["data"], dict) else row["data"]
            else:
                d = json.loads(row["data"])
            d["_bank_id"] = row[0]
            d["_subject"] = row[1]
            result.append(d)
        return result
    except Exception as e:
        print("db_load_bank error:", e)
        return []

def db_add_to_bank(subject, question):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if db_type == "postgres":
            import psycopg2.extras
            cur.execute("INSERT INTO question_bank (subject, data) VALUES (%s,%s)",
                       (subject, psycopg2.extras.Json(question)))
        else:
            cur.execute("INSERT INTO question_bank (subject, data, created_at) VALUES (?,?,?)",
                       (subject, json.dumps(question), datetime.datetime.now().isoformat()))
        conn.commit(); cur.close(); conn.close()
        return True
    except Exception as e:
        print("db_add_to_bank error:", e)
        return False

def db_delete_from_bank(bank_id):
    try:
        conn, db_type = get_db()
        cur = conn.cursor()
        if db_type == "postgres":
            cur.execute("DELETE FROM question_bank WHERE id=%s", (bank_id,))
        else:
            cur.execute("DELETE FROM question_bank WHERE id=?", (bank_id,))
        conn.commit(); cur.close(); conn.close()
        return True
    except Exception as e:
        print("db_delete_from_bank error:", e)
        return False
