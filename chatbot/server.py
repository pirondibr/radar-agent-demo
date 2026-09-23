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
from pipeline import STEP_DEFS, run_pipeline

STATIC_DIR = Path(__file__).resolve().parent / "static"
app = Flask(__name__, static_folder=str(STATIC_DIR), static_url_path="/static")

# Public demo host: only Chatguru sample data (no live scrapes)
DEMO_ONLY = os.environ.get("DEMO_ONLY", "").strip().lower() in ("1", "true", "yes")


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

        report = run_pipeline(job.parsed, _emit, _set_step)
        job.report = report
        job.status = "done"
        emit(job, "complete", report=report, progressive=True)
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
    if DEMO_ONLY:
        greeting = (
            "Ola! Eu sou o **Radar Agent** (demo publica **v1.0.9**). "
            "Aqui voce testa o fluxo completo com o exemplo **Chatguru**, "
            "sem scrapes ao vivo."
        )
        ask = (
            "Digite **demo** para ver Briefing, Concorrentes, Google Ads, SEO e Marca. "
            "Outras URLs ainda nao rodam neste ambiente (analise ao vivo vem na versao completa)."
        )
        examples = ["demo", "https://chatguru.com.br/"]
    else:
        greeting = (
            "Ola! Eu sou o **Radar Agent**, seu consultor de estrategia em marketing digital. "
            "Analiso a concorrencia e os canais de marketing da sua empresa."
        )
        ask = (
            "Para comecar, me diga a **empresa** (URL ou nome). "
            "Se quiser, indique de **1 a 3 concorrentes**. Nao e obrigatorio, "
            "a IA pesquisa automaticamente."
        )
        examples = [
            "https://chatguru.com.br/",
            "demo",
            "minhaempresa.com.br concorrentes: blip, huggy",
        ]
    return jsonify({
        "greeting": greeting,
        "ask": ask,
        "examples": examples,
        "steps": STEP_DEFS,
        "demo_only": DEMO_ONLY,
        "version": "1.0.9",
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
        # Public demo: only seeded Chatguru — never silently remap other URLs
        is_demo_request = bool(parsed.demo) or (parsed.slug or "").lower() == "chatguru"
        if not is_demo_request:
            return jsonify({
                "error": (
                    "Na demo publica so o exemplo **Chatguru** esta disponivel. "
                    "Digite **demo** para ver o fluxo. "
                    "Analise ao vivo de outros sites ainda nao roda neste ambiente."
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
    print(f"Radar Agent Chatbot -> http://127.0.0.1:{port}  DEMO_ONLY={DEMO_ONLY}")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
