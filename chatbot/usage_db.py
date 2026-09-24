# -*- coding: utf-8 -*-
"""SQLite: usos do chatbot + logs de execucao por etapa/artefato."""

from __future__ import annotations

import json
import os
import sqlite3
import threading
import time
from pathlib import Path
from typing import Any, Optional

DATA_DIR = Path(
    os.environ.get("RADAR_DATA_DIR", "").strip()
    or str(Path(__file__).resolve().parent / "data")
)
DB_PATH = DATA_DIR / "usage.sqlite"
RUNS_DIR = DATA_DIR / "runs"

_lock = threading.Lock()
_initialized = False


def _connect() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn


def init_db() -> None:
    global _initialized
    with _lock:
        if _initialized:
            return
        conn = _connect()
        try:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS runs (
                  id TEXT PRIMARY KEY,
                  kind TEXT NOT NULL DEFAULT 'free',
                  mode TEXT,
                  status TEXT NOT NULL DEFAULT 'pending',
                  company TEXT,
                  slug TEXT,
                  url TEXT,
                  demo INTEGER DEFAULT 0,
                  preferred_competitors_json TEXT,
                  parent_run_id TEXT,
                  started_at REAL,
                  finished_at REAL,
                  duration_ms INTEGER,
                  error_message TEXT,
                  report_summary_json TEXT,
                  user_agent TEXT,
                  ip TEXT
                );

                CREATE TABLE IF NOT EXISTS step_logs (
                  id INTEGER PRIMARY KEY AUTOINCREMENT,
                  run_id TEXT NOT NULL,
                  step_id TEXT NOT NULL,
                  step_index INTEGER,
                  state TEXT NOT NULL,
                  detail TEXT,
                  started_at REAL,
                  finished_at REAL,
                  duration_ms INTEGER,
                  companies_found INTEGER,
                  altos_found INTEGER,
                  medios_found INTEGER,
                  baixos_found INTEGER,
                  competitors_count_ui INTEGER,
                  display_tier TEXT,
                  error_message TEXT,
                  meta_json TEXT,
                  UNIQUE(run_id, step_id)
                );

                CREATE TABLE IF NOT EXISTS artifacts (
                  id INTEGER PRIMARY KEY AUTOINCREMENT,
                  run_id TEXT NOT NULL,
                  step_id TEXT,
                  kind TEXT NOT NULL,
                  path TEXT NOT NULL,
                  created_at REAL,
                  bytes INTEGER,
                  meta_json TEXT
                );

                CREATE TABLE IF NOT EXISTS event_logs (
                  id INTEGER PRIMARY KEY AUTOINCREMENT,
                  run_id TEXT NOT NULL,
                  ts REAL NOT NULL,
                  level TEXT DEFAULT 'info',
                  source TEXT DEFAULT 'pipeline',
                  message TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS leads (
                  id TEXT PRIMARY KEY,
                  ts REAL NOT NULL,
                  offer TEXT,
                  channel TEXT,
                  name TEXT,
                  contact_type TEXT,
                  contact TEXT,
                  company TEXT,
                  slug TEXT,
                  job_id TEXT,
                  order_id TEXT,
                  paid INTEGER DEFAULT 0,
                  status TEXT DEFAULT 'pending_manual',
                  meta_json TEXT
                );

                CREATE INDEX IF NOT EXISTS idx_runs_started ON runs(started_at DESC);
                CREATE INDEX IF NOT EXISTS idx_steps_run ON step_logs(run_id);
                CREATE INDEX IF NOT EXISTS idx_artifacts_run ON artifacts(run_id);
                CREATE INDEX IF NOT EXISTS idx_events_run ON event_logs(run_id, ts);
                CREATE INDEX IF NOT EXISTS idx_leads_ts ON leads(ts DESC);
                """
            )
            conn.commit()
            _initialized = True
        finally:
            conn.close()


def create_run(
    *,
    run_id: str,
    kind: str = "free",
    company: str = "",
    slug: str = "",
    url: str = "",
    demo: bool = False,
    preferred_competitors: Optional[list[str]] = None,
    parent_run_id: str = "",
    user_agent: str = "",
    ip: str = "",
) -> None:
    init_db()
    now = time.time()
    with _lock:
        conn = _connect()
        try:
            conn.execute(
                """
                INSERT OR REPLACE INTO runs (
                  id, kind, status, company, slug, url, demo,
                  preferred_competitors_json, parent_run_id,
                  started_at, user_agent, ip
                ) VALUES (?, ?, 'running', ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    kind or "free",
                    company or "",
                    slug or "",
                    url or "",
                    1 if demo else 0,
                    json.dumps(preferred_competitors or [], ensure_ascii=False),
                    parent_run_id or None,
                    now,
                    user_agent or "",
                    ip or "",
                ),
            )
            conn.commit()
        finally:
            conn.close()
    (RUNS_DIR / run_id).mkdir(parents=True, exist_ok=True)


def set_run_mode(run_id: str, mode: str) -> None:
    init_db()
    with _lock:
        conn = _connect()
        try:
            conn.execute("UPDATE runs SET mode = ? WHERE id = ?", (mode, run_id))
            conn.commit()
        finally:
            conn.close()


def finish_run(
    run_id: str,
    *,
    status: str,
    error_message: str = "",
    report_summary: Optional[dict[str, Any]] = None,
) -> None:
    init_db()
    now = time.time()
    with _lock:
        conn = _connect()
        try:
            row = conn.execute(
                "SELECT started_at FROM runs WHERE id = ?", (run_id,)
            ).fetchone()
            started = float(row["started_at"]) if row and row["started_at"] else now
            duration_ms = int(max(0, (now - started) * 1000))
            conn.execute(
                """
                UPDATE runs SET
                  status = ?, finished_at = ?, duration_ms = ?,
                  error_message = ?, report_summary_json = ?
                WHERE id = ?
                """,
                (
                    status,
                    now,
                    duration_ms,
                    error_message or "",
                    json.dumps(report_summary or {}, ensure_ascii=False),
                    run_id,
                ),
            )
            conn.commit()
        finally:
            conn.close()


def upsert_step(
    run_id: str,
    step_id: str,
    state: str,
    detail: str = "",
    *,
    step_index: Optional[int] = None,
    companies_found: Optional[int] = None,
    altos_found: Optional[int] = None,
    medios_found: Optional[int] = None,
    baixos_found: Optional[int] = None,
    competitors_count_ui: Optional[int] = None,
    display_tier: Optional[str] = None,
    error_message: str = "",
    meta: Optional[dict[str, Any]] = None,
) -> None:
    init_db()
    now = time.time()
    with _lock:
        conn = _connect()
        try:
            existing = conn.execute(
                "SELECT * FROM step_logs WHERE run_id = ? AND step_id = ?",
                (run_id, step_id),
            ).fetchone()
            if not existing:
                conn.execute(
                    """
                    INSERT INTO step_logs (
                      run_id, step_id, step_index, state, detail,
                      started_at, finished_at, duration_ms,
                      companies_found, altos_found, medios_found, baixos_found,
                      competitors_count_ui, display_tier, error_message, meta_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        run_id,
                        step_id,
                        step_index,
                        state,
                        detail,
                        now if state == "running" else None,
                        now if state in ("done", "error") else None,
                        0 if state in ("done", "error") else None,
                        companies_found,
                        altos_found,
                        medios_found,
                        baixos_found,
                        competitors_count_ui,
                        display_tier,
                        error_message or None,
                        json.dumps(meta, ensure_ascii=False) if meta else None,
                    ),
                )
            else:
                started = existing["started_at"]
                finished = existing["finished_at"]
                duration = existing["duration_ms"]
                if state == "running" and not started:
                    started = now
                if state in ("done", "error"):
                    finished = now
                    if started:
                        duration = int(max(0, (now - float(started)) * 1000))
                conn.execute(
                    """
                    UPDATE step_logs SET
                      step_index = COALESCE(?, step_index),
                      state = ?,
                      detail = ?,
                      started_at = ?,
                      finished_at = ?,
                      duration_ms = ?,
                      companies_found = COALESCE(?, companies_found),
                      altos_found = COALESCE(?, altos_found),
                      medios_found = COALESCE(?, medios_found),
                      baixos_found = COALESCE(?, baixos_found),
                      competitors_count_ui = COALESCE(?, competitors_count_ui),
                      display_tier = COALESCE(?, display_tier),
                      error_message = COALESCE(?, error_message),
                      meta_json = COALESCE(?, meta_json)
                    WHERE run_id = ? AND step_id = ?
                    """,
                    (
                        step_index,
                        state,
                        detail,
                        started,
                        finished,
                        duration,
                        companies_found,
                        altos_found,
                        medios_found,
                        baixos_found,
                        competitors_count_ui,
                        display_tier,
                        error_message or None,
                        json.dumps(meta, ensure_ascii=False) if meta else None,
                        run_id,
                        step_id,
                    ),
                )
            conn.commit()
        finally:
            conn.close()


def append_log(
    run_id: str,
    message: str,
    *,
    level: str = "info",
    source: str = "pipeline",
) -> None:
    if not message:
        return
    init_db()
    with _lock:
        conn = _connect()
        try:
            conn.execute(
                """
                INSERT INTO event_logs (run_id, ts, level, source, message)
                VALUES (?, ?, ?, ?, ?)
                """,
                (run_id, time.time(), level, source, message[:4000]),
            )
            conn.commit()
        finally:
            conn.close()
    # Also append to run log file for easy download
    try:
        log_path = RUNS_DIR / run_id / "pipeline.log"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%H:%M:%S')} [{level}] {message}\n")
    except OSError:
        pass


def add_artifact(
    run_id: str,
    *,
    kind: str,
    path: str,
    step_id: str = "",
    meta: Optional[dict[str, Any]] = None,
) -> None:
    init_db()
    p = Path(path) if path else None
    size = int(p.stat().st_size) if p and p.exists() else None
    with _lock:
        conn = _connect()
        try:
            conn.execute(
                """
                INSERT INTO artifacts (run_id, step_id, kind, path, created_at, bytes, meta_json)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    run_id,
                    step_id or None,
                    kind,
                    path,
                    time.time(),
                    size,
                    json.dumps(meta, ensure_ascii=False) if meta else None,
                ),
            )
            conn.commit()
        finally:
            conn.close()


def save_json_artifact(
    run_id: str,
    filename: str,
    data: Any,
    *,
    kind: str,
    step_id: str = "",
    meta: Optional[dict[str, Any]] = None,
) -> Path:
    dest = RUNS_DIR / run_id / filename
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    add_artifact(run_id, kind=kind, path=str(dest), step_id=step_id, meta=meta)
    return dest


def list_runs(limit: int = 50, offset: int = 0) -> list[dict[str, Any]]:
    init_db()
    with _lock:
        conn = _connect()
        try:
            rows = conn.execute(
                """
                SELECT * FROM runs
                ORDER BY COALESCE(started_at, 0) DESC
                LIMIT ? OFFSET ?
                """,
                (limit, offset),
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()


def get_run(run_id: str) -> Optional[dict[str, Any]]:
    init_db()
    with _lock:
        conn = _connect()
        try:
            row = conn.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
            return dict(row) if row else None
        finally:
            conn.close()


def get_run_steps(run_id: str) -> list[dict[str, Any]]:
    init_db()
    with _lock:
        conn = _connect()
        try:
            rows = conn.execute(
                """
                SELECT * FROM step_logs
                WHERE run_id = ?
                ORDER BY COALESCE(step_index, 999), id
                """,
                (run_id,),
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()


def get_run_artifacts(run_id: str) -> list[dict[str, Any]]:
    init_db()
    with _lock:
        conn = _connect()
        try:
            rows = conn.execute(
                "SELECT * FROM artifacts WHERE run_id = ? ORDER BY id",
                (run_id,),
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()


def get_run_logs(run_id: str, limit: int = 500) -> list[dict[str, Any]]:
    init_db()
    with _lock:
        conn = _connect()
        try:
            rows = conn.execute(
                """
                SELECT * FROM event_logs
                WHERE run_id = ?
                ORDER BY id DESC
                LIMIT ?
                """,
                (run_id, limit),
            ).fetchall()
            return [dict(r) for r in reversed(rows)]
        finally:
            conn.close()


def summarize_report(report: Optional[dict[str, Any]]) -> dict[str, Any]:
    if not report:
        return {}
    comps = report.get("competitors") or []
    return {
        "client": report.get("client"),
        "competitors_ui": len([c for c in comps if not c.get("is_client")]),
        "has_google_ads": bool(report.get("google_ads")),
        "has_seo": bool(report.get("seo")),
        "has_brand": bool(report.get("brand")),
        "competitors_note": report.get("competitors_note") or "",
        "display_tier": report.get("display_tier") or "",
    }


def save_lead(lead: dict[str, Any]) -> None:
    """Persiste lead Pro (contato pos-pagamento) no SQLite."""
    init_db()
    with _lock:
        conn = _connect()
        try:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS leads (
                  id TEXT PRIMARY KEY,
                  ts REAL NOT NULL,
                  offer TEXT,
                  channel TEXT,
                  name TEXT,
                  contact_type TEXT,
                  contact TEXT,
                  company TEXT,
                  slug TEXT,
                  job_id TEXT,
                  order_id TEXT,
                  paid INTEGER DEFAULT 0,
                  status TEXT DEFAULT 'pending_manual',
                  meta_json TEXT
                )
                """
            )
            try:
                conn.execute("ALTER TABLE leads ADD COLUMN name TEXT")
            except Exception:
                pass
            conn.execute(
                """
                INSERT OR REPLACE INTO leads (
                  id, ts, offer, channel, name, contact_type, contact,
                  company, slug, job_id, order_id, paid, status, meta_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    lead.get("id") or "",
                    float(lead.get("ts") or time.time()),
                    lead.get("offer") or "",
                    lead.get("channel") or "",
                    lead.get("name") or "",
                    lead.get("contact_type") or "",
                    lead.get("contact") or "",
                    lead.get("company") or "",
                    lead.get("slug") or "",
                    lead.get("job_id") or "",
                    lead.get("order_id") or "",
                    1 if lead.get("paid") else 0,
                    lead.get("status") or "pending_manual",
                    json.dumps(
                        {k: v for k, v in lead.items() if k not in {
                            "id", "ts", "offer", "channel", "name", "contact_type", "contact",
                            "company", "slug", "job_id", "order_id", "paid", "status",
                        }},
                        ensure_ascii=False,
                    ),
                ),
            )
            conn.commit()
        finally:
            conn.close()


def list_leads(limit: int = 100) -> list[dict[str, Any]]:
    init_db()
    with _lock:
        conn = _connect()
        try:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS leads (
                  id TEXT PRIMARY KEY,
                  ts REAL NOT NULL,
                  offer TEXT,
                  channel TEXT,
                  contact_type TEXT,
                  contact TEXT,
                  company TEXT,
                  slug TEXT,
                  job_id TEXT,
                  order_id TEXT,
                  paid INTEGER DEFAULT 0,
                  status TEXT DEFAULT 'pending_manual',
                  meta_json TEXT
                )
                """
            )
            rows = conn.execute(
                "SELECT * FROM leads ORDER BY ts DESC LIMIT ?",
                (limit,),
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()
