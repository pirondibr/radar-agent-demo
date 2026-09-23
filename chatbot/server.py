# -*- coding: utf-8 -*-
"""
Radar Agent Chatbot | MVP SEO + Google Ads.

Uso:
    python server.py
    Abra http://127.0.0.1:8766
"""

from __future__ import annotations

import json
import os
import threading
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from queue import Empty, Queue
from typing import Any, Optional

from flask import Flask, Response, jsonify, request, send_from_directory

from parse_input import ParsedInput, parse_user_message
from pipeline import EXTRA_STEP_DEFS, STEP_DEFS, run_extras_pipeline, run_pipeline

STATIC_DIR = Path(__file__).resolve().parent / "static"
app = Flask(__name__, static_folder=str(STATIC_DIR), static_url_path="/static")


def _load_dotenv() -> None:
    """Load Radar/.env into os.environ if present (local + optional Render secret file)."""
    env_path = Path(__file__).resolve().parent.parent / ".env"
    if not env_path.exists():
        return
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key = key.strip()
        val = val.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = val


_load_dotenv()

# Public sample-only host when DEMO_ONLY=1. Live needs API keys (see .env.example).
DEMO_ONLY = os.environ.get("DEMO_ONLY", "").strip().lower() in ("1", "true", "yes")
APP_VERSION = "1.2.0"

REQUIRED_LIVE_KEYS = (
    "OPENROUTER_API_KEY",
    "SCRAPINGBEE_API_KEY",
    "DATAFORSEO_USER",
    "DATAFORSEO_PASS",
    "SEMRUSH_API_KEY",
)

LEADS_DIR = Path(__file__).resolve().parent / "data"
LEADS_FILE = LEADS_DIR / "leads.jsonl"


def _live_keys_ready() -> dict[str, bool]:
    return {k: bool(os.environ.get(k, "").strip()) for k in REQUIRED_LIVE_KEYS}


def _missing_live_keys() -> list[str]:
    return [k for k, ok in _live_keys_ready().items() if not ok]


@dataclass
class Job:
    job_id: str
    parsed: ParsedInput
    queue: Queue = field(default_factory=Queue)
    status: str = "pending"
    current_step: str = ""
    steps: dict = field(default_factory=dict)
    error: Optional[str] = None
    report: Optional[dict] = None


jobs: dict[str, Job] = {}


def emit(job: Job, event_type: str, **payload: Any) -> None:
    msg = {"type": event_type, "job_id": job.job_id, "ts": time.time(), **payload}
    job.queue.put(msg)


def set_step(job: Job, step_id: str, state: str, detail: str = "") -> None:
    job.current_step = step_id
    job.steps[step_id] = {"state": state, "detail": detail, "at": time.time()}
    emit(job, "step", step_id=step_id, state=state, detail=detail, steps=job.steps)


def _worker(job: Job) -> None:
    try:
        job.status = "running"
        emit(job, "started", company=job.parsed.company, slug=job.parsed.slug, demo=job.parsed.demo)

        def _emit(event_type: str, **payload: Any) -> None:
            emit(job, event_type, **payload)

        def _set_step(step_id: str, state: str, detail: str = "") -> None:
            set_step(job, step_id, state, detail)

        kind = getattr(job.parsed, "kind", "") or ""
        if kind == "extras":
            report = run_extras_pipeline(
                slug=job.parsed.slug or "chatguru",
                company=job.parsed.company or "",
                emit=_emit,
                set_step=_set_step,
                demo=bool(job.parsed.demo),
                preferred_competitors=job.parsed.competitors,
            )
        else:
            report = run_pipeline(job.parsed, _emit, _set_step)
        job.report = report
        job.status = "done"
        emit(job, "complete", report=report, progressive=True, extras=(kind == "extras"))
    except Exception as e:
        job.status = "error"
        job.error = str(e)
        if job.current_step:
            job.steps[job.current_step] = {
                "state": "error",
                "detail": str(e),
                "at": time.time(),
            }
        emit(job, "error", message=str(e), steps=job.steps)


@app.get("/")
def index():
    resp = send_from_directory(STATIC_DIR, "index.html")
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    return resp


@app.get("/api/hello")
def hello():
    keys = _live_keys_ready()
    live_ok = (not DEMO_ONLY) and all(keys.values())
    if DEMO_ONLY:
        greeting = (
            f"Ola! Eu sou o **Radar Agent** (demo publica **v{APP_VERSION}**). "
            "Aqui voce testa o fluxo completo com o exemplo **Chatguru**, "
            "sem scrapes ao vivo."
        )
        ask = (
            "Digite **demo** para ver Briefing, Concorrentes, Google Ads, SEO e Marca. "
            "Outras URLs ainda nao rodam neste ambiente."
        )
        examples = ["demo", "https://chatguru.com.br/"]
    else:
        greeting = (
            f"Ola! Eu sou o **Radar Agent** (**v{APP_VERSION}**). "
            "Analiso a concorrencia e os canais de marketing da sua empresa ao vivo."
        )
        if live_ok:
            ask = (
                "Envie a **URL do site** (ex: https://www.mendesortega.com.br/). "
                "Opcional: 1 a 3 concorrentes. Digite **demo** para o exemplo Chatguru (sem scrape)."
            )
        else:
            missing = ", ".join(_missing_live_keys())
            ask = (
                "Live ainda sem todas as API keys no servidor "
                f"(faltam: {missing}). Digite **demo** para o exemplo Chatguru, "
                "ou configure as keys no Render Environment."
            )
        examples = [
            "https://www.mendesortega.com.br/",
            "demo",
            "https://chatguru.com.br/",
        ]
    return jsonify({
        "greeting": greeting,
        "ask": ask,
        "examples": examples,
        "steps": STEP_DEFS,
        "demo_only": DEMO_ONLY,
        "live_ready": live_ok,
        "keys_ready": keys,
        "version": APP_VERSION,
    })


@app.post("/api/chat")
def chat():
    data = request.get_json(silent=True) or {}
    message = (data.get("message") or "").strip()
    if not message:
        return jsonify({"error": "Mensagem vazia"}), 400

    parsed = parse_user_message(message)
    if not parsed.company and not parsed.slug and not parsed.demo:
        return jsonify({
            "error": "Nao consegui identificar a empresa. Envie uma URL ou o nome (ex: https://chatguru.com.br/).",
            "parsed": None,
        }), 400

    if DEMO_ONLY:
        # Public sample-only: never silently remap other URLs
        is_demo_request = bool(parsed.demo) or (parsed.slug or "").lower() == "chatguru"
        if not is_demo_request:
            return jsonify({
                "error": (
                    "Neste modo so o exemplo **Chatguru** esta disponivel. "
                    "Digite **demo** ou configure DEMO_ONLY=0 + API keys para live."
                ),
                "parsed": {
                    "company": parsed.company,
                    "url": parsed.url,
                    "slug": parsed.slug,
                    "competitors": parsed.competitors,
                    "demo": False,
                },
                "demo_only": True,
            }), 400
        parsed.demo = True
        parsed.company = "Chatguru (demo)"
        parsed.slug = "chatguru"
        parsed.url = parsed.url or "https://chatguru.com.br/"
    elif not parsed.demo:
        missing = _missing_live_keys()
        if missing:
            return jsonify({
                "error": (
                    "Analise ao vivo precisa das API keys no Render Environment: "
                    + ", ".join(missing)
                    + ". Enquanto isso, digite **demo** para o exemplo Chatguru."
                ),
                "missing_keys": missing,
                "demo_only": False,
            }), 503

    job_id = uuid.uuid4().hex[:12]
    job = Job(job_id=job_id, parsed=parsed)
    for s in STEP_DEFS:
        job.steps[s["id"]] = {"state": "pending", "detail": "", "at": 0}
    jobs[job_id] = job

    threading.Thread(target=_worker, args=(job,), daemon=True).start()

    return jsonify({
        "job_id": job_id,
        "parsed": {
            "company": parsed.company,
            "url": parsed.url,
            "slug": parsed.slug,
            "competitors": parsed.competitors,
            "demo": parsed.demo,
        },
        "steps": STEP_DEFS,
        "ack": _ack_message(parsed),
        "demo_only": DEMO_ONLY,
    })


def _ack_message(parsed: ParsedInput) -> str:
    comps = ""
    if parsed.competitors:
        comps = " Concorrentes sugeridos: " + ", ".join(parsed.competitors) + "."
    else:
        comps = " Sem concorrentes informados. Vou pesquisar automaticamente."
    if DEMO_ONLY or parsed.demo:
        mode = " (demo publica, resultados em etapas, cerca de 40s)"
    else:
        mode = " (pipeline completa, tempo estimado cerca de 10 min. Resultados aparecem etapa a etapa)"
    return (
        f"Perfeito. Vou analisar **{parsed.company or parsed.slug}**{mode}.{comps} "
        "Vou liberar Briefing, Concorrentes, Google Ads, SEO e Marca conforme cada etapa terminar."
    )


@app.post("/api/leads")
def create_lead():
    """MVP pago simulado: captura email/WhatsApp para analise profunda manual."""
    data = request.get_json(silent=True) or {}
    contact_type = (data.get("contact_type") or "").strip().lower()
    contact = (data.get("contact") or "").strip()
    channel = (data.get("channel") or "").strip()
    if contact_type not in ("email", "whatsapp"):
        return jsonify({"error": "Informe contact_type: email ou whatsapp"}), 400
    if not contact or len(contact) < 5:
        return jsonify({"error": "Informe um contato valido"}), 400

    lead = {
        "id": uuid.uuid4().hex[:12],
        "ts": time.time(),
        "offer": data.get("offer") or "deep_channel",
        "channel": channel,
        "contact_type": contact_type,
        "contact": contact,
        "company": (data.get("company") or "").strip(),
        "slug": (data.get("slug") or "").strip(),
        "job_id": (data.get("job_id") or "").strip(),
        "simulated_payment": True,
        "status": "pending_manual",
    }
    LEADS_DIR.mkdir(parents=True, exist_ok=True)
    with LEADS_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(lead, ensure_ascii=False) + "\n")
    print(f"[LEAD] {lead['id']} {lead['channel']} {lead['contact_type']}={lead['contact']} slug={lead['slug']}")
    return jsonify({
        "ok": True,
        "lead_id": lead["id"],
        "message": (
            f"Perfeito. Sua analise Pro de **{channel or 'canal'}** esta na fila. "
            f"Em ate 24h enviamos no seu {contact_type}."
        ),
    })


@app.post("/api/extras")
def start_extras():
    """Canais extra: Meta, LinkedIn, Instagram, YouTube (sem TikTok)."""
    data = request.get_json(silent=True) or {}
    slug = (data.get("slug") or "").strip() or "chatguru"
    company = (data.get("company") or "").strip() or slug
    demo = bool(data.get("demo"))
    job_id_src = (data.get("job_id") or "").strip()

    if DEMO_ONLY:
        demo = True
        slug = "chatguru"
        company = company or "Chatguru (demo)"
    elif not demo:
        missing = _missing_live_keys()
        # extras precisam sobretudo ScrapingBee
        if "SCRAPINGBEE_API_KEY" in missing:
            return jsonify({
                "error": "Canais extra precisam de SCRAPINGBEE_API_KEY no servidor.",
                "missing_keys": missing,
            }), 503

    parsed = ParsedInput(
        company=company,
        slug=slug,
        demo=demo,
        kind="extras",
        raw=f"extras:{slug}",
    )
    job_id = uuid.uuid4().hex[:12]
    job = Job(job_id=job_id, parsed=parsed)
    for s in EXTRA_STEP_DEFS:
        job.steps[s["id"]] = {"state": "pending", "detail": "", "at": 0}
    jobs[job_id] = job
    threading.Thread(target=_worker, args=(job,), daemon=True).start()

    return jsonify({
        "job_id": job_id,
        "parent_job_id": job_id_src or None,
        "parsed": {
            "company": company,
            "slug": slug,
            "demo": demo,
            "kind": "extras",
        },
        "steps": EXTRA_STEP_DEFS,
        "ack": (
            f"Vou analisar os **canais extra** de **{company}**: Meta Ads, LinkedIn, Instagram e YouTube "
            "(TikTok em breve). Resultados aparecem etapa a etapa."
        ),
    })


@app.get("/api/jobs/<job_id>/events")
def job_events(job_id: str):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job nao encontrado"}), 404

    def stream():
        # Snapshot inicial
        yield _sse({"type": "snapshot", "job_id": job_id, "status": job.status, "steps": job.steps})
        while True:
            try:
                msg = job.queue.get(timeout=15)
            except Empty:
                yield _sse({"type": "ping", "job_id": job_id, "status": job.status})
                if job.status in ("done", "error"):
                    break
                continue
            yield _sse(msg)
            if msg.get("type") in ("complete", "error"):
                break

    return Response(
        stream(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )


@app.get("/api/jobs/<job_id>")
def job_status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job nao encontrado"}), 404
    return jsonify({
        "job_id": job_id,
        "status": job.status,
        "steps": job.steps,
        "error": job.error,
        "report": job.report,
        "parsed": {
            "company": job.parsed.company,
            "url": job.parsed.url,
            "slug": job.parsed.slug,
            "competitors": job.parsed.competitors,
            "demo": job.parsed.demo,
        },
    })


def _sse(payload: dict) -> str:
    return f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8766"))
    print(f"Radar Agent Chatbot -> http://127.0.0.1:{port}  DEMO_ONLY={DEMO_ONLY} live_keys={_live_keys_ready()}")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
