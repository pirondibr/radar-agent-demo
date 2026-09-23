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
from mercadopago_client import (
    EXTRAS_PRICE,
    PRO_PRICE,
    apply_payment_to_order,
    create_pro_checkout,
    fetch_payment,
    load_order,
    mark_order_contact,
    mp_configured,
    mp_public_key,
)

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
APP_VERSION = "1.3.2"

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
    greeting = (
        "Olá! Eu sou o Agente de IA **Radar da Concorrência**. "
        "Eu posso analisar os canais de marketing da sua empresa e dos "
        "seus concorrentes, e te mostrar onde você está ganhando ou perdendo "
        "espaço e dinheiro no seu nicho."
    )
    ask = "Para iniciarmos, envie o endereço do seu site logo abaixo."
    if DEMO_ONLY:
        examples = ["https://www.mendesortega.com.br/"]
    elif live_ok:
        examples = ["https://www.mendesortega.com.br/"]
    else:
        missing = ", ".join(_missing_live_keys())
        ask = (
            "Para iniciarmos, envie o endereço do seu site logo abaixo. "
            f"(Servidor ainda sem todas as API keys: {missing}.)"
        )
        examples = ["https://www.mendesortega.com.br/"]
    return jsonify({
        "greeting": greeting,
        "ask": ask,
        "examples": examples,
        "steps": STEP_DEFS,
        "demo_only": DEMO_ONLY,
        "live_ready": live_ok,
        "keys_ready": keys,
        "payments_ready": mp_configured(),
        "pro_price": PRO_PRICE,
        "extras_price": EXTRAS_PRICE,
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
                    "Neste ambiente a analise ao vivo ainda nao esta liberada. "
                    "Configure DEMO_ONLY=0 + API keys para analisar o seu site."
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
                    + "."
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
    name = parsed.company or parsed.slug or "seu site"
    if parsed.competitors:
        return (
            f"Perfeito. Vou analisar seu site **{name}** e os concorrentes "
            f"informados ({', '.join(parsed.competitors)}). "
            "(tempo estimado cerca de 5 min.)."
        )
    return (
        f"Perfeito. Vou analisar seu site **{name}** e encontrar seus concorrentes. "
        "(tempo estimado cerca de 5 min.)."
    )


@app.post("/api/checkout")
def checkout():
    """Cria Checkout Pro (PIX + cartao) para Radar Pro R$99."""
    data = request.get_json(silent=True) or {}
    channel = (data.get("channel") or "").strip() or "canal"
    if not mp_configured():
        return jsonify({
            "error": (
                "Mercado Pago ainda nao esta configurado. "
                "Defina MERCADOPAGO_ACCESS_TOKEN (e PUBLIC_BASE_URL) no ambiente."
            ),
            "payments_ready": False,
        }), 503
    try:
        result = create_pro_checkout(
            channel=channel,
            company=(data.get("company") or "").strip(),
            slug=(data.get("slug") or "").strip(),
            job_id=(data.get("job_id") or "").strip(),
            product=(data.get("product") or "deep_channel").strip().lower(),
        )
        return jsonify({"ok": True, **result})
    except Exception as e:
        return jsonify({"error": str(e)}), 502


@app.get("/api/checkout/<order_id>")
def checkout_status(order_id: str):
    order = load_order(order_id)
    if not order:
        return jsonify({"error": "Pedido nao encontrado"}), 404
    return jsonify({
        "order_id": order["id"],
        "status": order.get("status"),
        "mp_status": order.get("mp_status"),
        "channel": order.get("channel"),
        "amount": order.get("amount"),
        "paid": order.get("status") == "paid",
        "preference_id": order.get("preference_id"),
        "init_point": order.get("init_point"),
    })


@app.post("/api/webhooks/mercadopago")
@app.get("/api/webhooks/mercadopago")
def mercadopago_webhook():
    """Notificacoes Mercado Pago (payment)."""
    payload = request.get_json(silent=True) or {}
    topic = (
        request.args.get("topic")
        or request.args.get("type")
        or payload.get("type")
        or payload.get("topic")
        or ""
    ).lower()
    payment_id = (
        request.args.get("data.id")
        or request.args.get("id")
        or str((payload.get("data") or {}).get("id") or "")
        or str(payload.get("id") or "")
    )
    print(f"[MP-WEBHOOK] topic={topic} payment_id={payment_id} args={dict(request.args)}")
    if "payment" in topic and payment_id and payment_id not in ("", "null"):
        try:
            payment = fetch_payment(str(payment_id))
            order = apply_payment_to_order(payment)
            print(f"[MP-WEBHOOK] order={order and order.get('id')} status={order and order.get('status')}")
        except Exception as e:
            print(f"[MP-WEBHOOK] erro: {e}")
            return jsonify({"ok": False, "error": str(e)}), 200
    return jsonify({"ok": True})


@app.get("/pay/return")
def pay_return():
    """Volta do Checkout Pro para o chat com order_id na query."""
    status = (request.args.get("status") or "").strip()
    order_id = (request.args.get("order_id") or "").strip()
    # Se MP mandar collection_id / payment_id, tenta sincronizar
    payment_id = (
        request.args.get("payment_id")
        or request.args.get("collection_id")
        or ""
    ).strip()
    if payment_id and payment_id not in ("null", "None") and mp_configured():
        try:
            apply_payment_to_order(fetch_payment(payment_id))
        except Exception as e:
            print(f"[MP-RETURN] sync erro: {e}")
    # Pagina minima que fecha o popup / redireciona ao app
    html = f"""<!doctype html><html lang="pt-BR"><head>
<meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Pagamento Radar Pro</title>
<style>body{{font-family:system-ui,sans-serif;background:#fafafa;color:#111;display:flex;min-height:100vh;align-items:center;justify-content:center;margin:0}}
.card{{background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:28px 24px;max-width:420px;text-align:center;box-shadow:0 8px 30px rgba(0,0,0,.06)}}
h1{{font-size:18px;margin:0 0 8px}}p{{color:#555;font-size:14px;line-height:1.5}}
a{{display:inline-block;margin-top:14px;padding:10px 16px;border-radius:10px;background:#4f46e5;color:#fff;text-decoration:none;font-weight:700}}</style>
</head><body><div class="card">
<h1>{"Pagamento aprovado" if status=="success" else ("Pagamento pendente" if status=="pending" else "Pagamento nao concluido")}</h1>
<p>{"Pagamento recebido. Volte ao chat para informar o contato e receber a analise." if status=="success" else "Se pagou via PIX, aguarde a confirmacao e volte ao chat."}</p>
<a href="/?paid={status}&order_id={order_id}">Voltar ao Radar da Concorrência</a>
</div>
<script>
try {{
  if (window.opener) {{
    window.opener.postMessage({{ type: 'mp_return', status: '{status}', order_id: '{order_id}' }}, '*');
  }}
}} catch (e) {{}}
</script>
</body></html>"""
    return Response(html, mimetype="text/html")


@app.post("/api/leads")
def create_lead():
    """Apos pagamento: captura email/WhatsApp para analise profunda."""
    data = request.get_json(silent=True) or {}
    contact_type = (data.get("contact_type") or "").strip().lower()
    contact = (data.get("contact") or "").strip()
    channel = (data.get("channel") or "").strip()
    order_id = (data.get("order_id") or "").strip()
    if contact_type not in ("email", "whatsapp"):
        return jsonify({"error": "Informe contact_type: email ou whatsapp"}), 400
    if not contact or len(contact) < 5:
        return jsonify({"error": "Informe um contato valido"}), 400

    paid = False
    if mp_configured():
        if not order_id:
            return jsonify({"error": "Pagamento obrigatorio. Conclua o checkout Mercado Pago primeiro."}), 402
        order = load_order(order_id)
        if not order:
            return jsonify({"error": "Pedido nao encontrado"}), 404
        if order.get("status") != "paid":
            return jsonify({
                "error": "Pagamento ainda nao confirmado. Conclua PIX/cartao e tente de novo.",
                "order_status": order.get("status"),
            }), 402
        paid = True
        mark_order_contact(order_id, contact_type, contact)
    elif order_id:
        order = load_order(order_id)
        if order and order.get("status") == "paid":
            paid = True
            mark_order_contact(order_id, contact_type, contact)

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
        "order_id": order_id,
        "simulated_payment": not paid,
        "status": "pending_manual",
        "paid": paid,
    }
    LEADS_DIR.mkdir(parents=True, exist_ok=True)
    with LEADS_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(lead, ensure_ascii=False) + "\n")
    print(f"[LEAD] {lead['id']} paid={paid} {lead['channel']} {lead['contact_type']}={lead['contact']}")
    return jsonify({
        "ok": True,
        "lead_id": lead["id"],
        "paid": paid,
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
    print(f"Radar da Concorrencia -> http://127.0.0.1:{port}  DEMO_ONLY={DEMO_ONLY} live_keys={_live_keys_ready()}")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
