import json
import os
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

DB_PATH = Path("data/gt35_v3.db")

def _secret(name, default=""):
    try:
        return st.secrets.get(name, os.getenv(name, default))
    except Exception:
        return os.getenv(name, default)

def is_supabase_configured():
    return bool(_secret("SUPABASE_URL") and _secret("SUPABASE_KEY"))

@st.cache_resource
def get_supabase():
    if not is_supabase_configured():
        return None
    from supabase import create_client
    return create_client(_secret("SUPABASE_URL"), _secret("SUPABASE_KEY"))

def init_sqlite():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.execute("""
    CREATE TABLE IF NOT EXISTS users (
        email TEXT PRIMARY KEY,
        password TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'user',
        farm TEXT DEFAULT 'ALL',
        active INTEGER NOT NULL DEFAULT 1
    )
    """)
    con.execute("""
    CREATE TABLE IF NOT EXISTS farms (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        region TEXT,
        name TEXT UNIQUE NOT NULL,
        capacity REAL DEFAULT 0,
        manager TEXT,
        active INTEGER NOT NULL DEFAULT 1,
        updated_at TEXT
    )
    """)
    con.execute("""
    CREATE TABLE IF NOT EXISTS gt35_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER NOT NULL,
        month INTEGER,
        week TEXT NOT NULL,
        region TEXT,
        farm TEXT NOT NULL,
        payload TEXT NOT NULL,
        status TEXT DEFAULT 'Nháp',
        created_by TEXT,
        created_at TEXT,
        updated_at TEXT,
        UNIQUE(year, week, farm)
    )
    """)
    con.execute("""
    INSERT OR IGNORE INTO users(email,password,role,farm,active)
    VALUES ('admin@gt35.local','admin123','admin','ALL',1)
    """)
    con.commit()
    con.close()

def seed_farms(initial_farms):
    if is_supabase_configured():
        try:
            sb = get_supabase()
            existing = sb.table("farms").select("name").execute().data or []
            existing_names = {x["name"] for x in existing}
            rows = [x for x in initial_farms if x["name"] not in existing_names]
            if rows:
                sb.table("farms").insert(rows).execute()
        except Exception:
            pass
        return
    init_sqlite()
    con = sqlite3.connect(DB_PATH)
    for f in initial_farms:
        con.execute(
            """INSERT OR IGNORE INTO farms(region,name,capacity,manager,active,updated_at)
               VALUES (?,?,?,?,?,?)""",
            (f.get("region",""), f["name"], f.get("capacity",0), f.get("manager",""),
             1 if f.get("active",True) else 0, datetime.now().isoformat(timespec="seconds"))
        )
    con.commit()
    con.close()

def login_user(email, password):
    if is_supabase_configured():
        try:
            sb = get_supabase()
            result = sb.auth.sign_in_with_password({"email": email, "password": password})
            profile = sb.table("profiles").select("*").eq("id", str(result.user.id)).single().execute()
            p = profile.data or {}
            st.session_state["user"] = {
                "id": str(result.user.id), "email": email,
                "role": p.get("role","user"), "farm": p.get("farm","ALL")
            }
            return True, "Đăng nhập thành công."
        except Exception as e:
            return False, f"Đăng nhập không thành công: {e}"
    init_sqlite()
    con = sqlite3.connect(DB_PATH)
    row = con.execute(
        "SELECT email,role,farm FROM users WHERE email=? AND password=? AND active=1",
        (email,password)
    ).fetchone()
    con.close()
    if not row:
        return False, "Sai email, mật khẩu hoặc tài khoản đã bị khóa."
    st.session_state["user"] = {"email":row[0], "role":row[1], "farm":row[2]}
    return True, "Đăng nhập thành công."

def logout_user():
    if is_supabase_configured():
        try:
            get_supabase().auth.sign_out()
        except Exception:
            pass
    st.session_state.pop("user", None)

def get_current_user():
    return st.session_state.get("user")

def load_farms(include_inactive=False):
    if is_supabase_configured():
        try:
            q = get_supabase().table("farms").select("*").order("name")
            if not include_inactive:
                q = q.eq("active", True)
            return pd.DataFrame(q.execute().data or [])
        except Exception:
            return pd.DataFrame()
    init_sqlite()
    con = sqlite3.connect(DB_PATH)
    sql = "SELECT id,region,name,capacity,manager,active,updated_at FROM farms"
    if not include_inactive:
        sql += " WHERE active=1"
    sql += " ORDER BY name"
    df = pd.read_sql_query(sql, con)
    con.close()
    if not df.empty:
        df["active"] = df["active"].astype(bool)
    return df

def save_farms(df):
    df = df.copy()
    expected = ["id","region","name","capacity","manager","active"]
    for c in expected:
        if c not in df.columns:
            df[c] = None
    df = df[df["name"].notna() & (df["name"].astype(str).str.strip() != "")]
    now = datetime.now().isoformat(timespec="seconds")
    if is_supabase_configured():
        try:
            sb = get_supabase()
            rows = []
            for _, r in df.iterrows():
                item = {
                    "region": str(r.get("region") or ""),
                    "name": str(r["name"]).strip(),
                    "capacity": float(r.get("capacity") or 0),
                    "manager": str(r.get("manager") or ""),
                    "active": bool(r.get("active", True)),
                    "updated_at": now,
                }
                rid = r.get("id")
                if pd.notna(rid):
                    item["id"] = int(rid)
                rows.append(item)
            if rows:
                sb.table("farms").upsert(rows, on_conflict="name").execute()
            return True, f"Đã lưu {len(rows)} trại."
        except Exception as e:
            return False, f"Lỗi Supabase: {e}"
    init_sqlite()
    try:
        con = sqlite3.connect(DB_PATH)
        for _, r in df.iterrows():
            con.execute("""
                INSERT INTO farms(region,name,capacity,manager,active,updated_at)
                VALUES (?,?,?,?,?,?)
                ON CONFLICT(name) DO UPDATE SET
                  region=excluded.region,
                  capacity=excluded.capacity,
                  manager=excluded.manager,
                  active=excluded.active,
                  updated_at=excluded.updated_at
            """, (
                str(r.get("region") or ""), str(r["name"]).strip(),
                float(r.get("capacity") or 0), str(r.get("manager") or ""),
                1 if bool(r.get("active",True)) else 0, now
            ))
        con.commit()
        con.close()
        return True, f"Đã lưu {len(df)} trại."
    except Exception as e:
        return False, f"Lỗi lưu trại: {e}"

def _clean_payload(payload):
    clean = {}
    for k,v in payload.items():
        if pd.isna(v) if not isinstance(v,(list,dict)) else False:
            clean[k] = None
        elif hasattr(v, "isoformat"):
            clean[k] = v.isoformat()
        elif isinstance(v, (int,float,str,bool)) or v is None:
            clean[k] = v
        else:
            clean[k] = str(v)
    return clean

def save_record(meta, payload, user_email, status="Nháp"):
    payload = _clean_payload(payload)
    now = datetime.now().isoformat(timespec="seconds")
    row = {
        "year": int(meta["year"]), "month": int(meta.get("month") or 0),
        "week": str(meta["week"]), "region": str(meta.get("region") or ""),
        "farm": str(meta["farm"]), "payload": payload,
        "status": status, "created_by": user_email,
        "created_at": now, "updated_at": now,
    }
    if is_supabase_configured():
        try:
            get_supabase().table("gt35_records").upsert(
                row, on_conflict="year,week,farm"
            ).execute()
            return True, "Đã lưu dữ liệu."
        except Exception as e:
            return False, f"Lỗi Supabase: {e}"
    init_sqlite()
    try:
        con = sqlite3.connect(DB_PATH)
        con.execute("""
        INSERT INTO gt35_records(year,month,week,region,farm,payload,status,created_by,created_at,updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(year,week,farm) DO UPDATE SET
          month=excluded.month, region=excluded.region, payload=excluded.payload,
          status=excluded.status, updated_at=excluded.updated_at
        """, (
            row["year"],row["month"],row["week"],row["region"],row["farm"],
            json.dumps(payload,ensure_ascii=False),status,user_email,now,now
        ))
        con.commit()
        con.close()
        return True, "Đã lưu dữ liệu."
    except Exception as e:
        return False, f"Lỗi lưu dữ liệu: {e}"

def load_records():
    if is_supabase_configured():
        try:
            rows = get_supabase().table("gt35_records").select("*").order("created_at", desc=True).execute().data or []
        except Exception:
            return pd.DataFrame()
    else:
        init_sqlite()
        con = sqlite3.connect(DB_PATH)
        rows = pd.read_sql_query("SELECT * FROM gt35_records ORDER BY id DESC", con).to_dict("records")
        con.close()
    flat = []
    for r in rows:
        p = r.get("payload") or {}
        if isinstance(p,str):
            try: p = json.loads(p)
            except Exception: p = {}
        item = {k:v for k,v in r.items() if k != "payload"}
        item.update(p)
        flat.append(item)
    return pd.DataFrame(flat)

def delete_record(record_id):
    if is_supabase_configured():
        try:
            get_supabase().table("gt35_records").delete().eq("id", int(record_id)).execute()
            return True, "Đã xóa bản ghi."
        except Exception as e:
            return False, str(e)
    init_sqlite()
    try:
        con=sqlite3.connect(DB_PATH)
        con.execute("DELETE FROM gt35_records WHERE id=?",(int(record_id),))
        con.commit(); con.close()
        return True,"Đã xóa bản ghi."
    except Exception as e:
        return False,str(e)
