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
from pipeline import (
    EXTRA_STEP_DEFS,
    STEP_DEFS,
    run_enrich_competitors,
    run_extras_pipeline,
    run_pipeline,
)
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
from meta_capi import capi_configured, send_capi_event, track_lead_from_request
import usage_db

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
APP_VERSION = "1.5.19"

try:
    usage_db.init_db()
except Exception as _e:
    print(f"[usage_db] init falhou: {_e}")

REQUIRED_LIVE_KEYS = (
    "OPENROUTER_API_KEY",
    "SCRAPINGBEE_API_KEY",
    "DATAFORSEO_USER",
    "DATAFORSEO_PASS",
    "SEMRUSH_API_KEY",
)

LEADS_DIR = Path(
    os.environ.get("RADAR_DATA_DIR", "").strip()
    or str(Path(__file__).resolve().parent / "data")
)
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
    awaiting: str = ""
    gate: threading.Event = field(default_factory=threading.Event)
    gate_payload: dict = field(default_factory=dict)


jobs: dict[str, Job] = {}


def emit(job: Job, event_type: str, **payload: Any) -> None:
    msg = {"type": event_type, "job_id": job.job_id, "ts": time.time(), **payload}
    job.queue.put(msg)
    try:
        if event_type == "log":
            usage_db.append_log(job.job_id, str(payload.get("line") or ""))
        elif event_type == "pipeline_meta":
            usage_db.set_run_mode(job.job_id, str(payload.get("mode") or ""))
        elif event_type == "error":
            usage_db.append_log(
                job.job_id,
                str(payload.get("message") or "erro"),
                level="error",
                source="worker",
            )
    except Exception:
        pass


def set_step(job: Job, step_id: str, state: str, detail: str = "") -> None:
    job.current_step = step_id
    job.steps[step_id] = {"state": state, "detail": detail, "at": time.time()}
    emit(job, "step", step_id=step_id, state=state, detail=detail, steps=job.steps)
    try:
        meta = next((s for s in (STEP_DEFS + EXTRA_STEP_DEFS) if s["id"] == step_id), None)
        usage_db.upsert_step(
            job.job_id,
            step_id,
            state,
            detail,
            step_index=int(meta["index"]) if meta else None,
            error_message=detail if state == "error" else "",
        )
    except Exception:
        pass


def _worker(job: Job) -> None:
    try:
        job.status = "running"
        emit(job, "started", company=job.parsed.company, slug=job.parsed.slug, demo=job.parsed.demo)

        def _emit(event_type: str, **payload: Any) -> None:
            emit(job, event_type, **payload)

        def _set_step(step_id: str, state: str, detail: str = "") -> None:
            set_step(job, step_id, state, detail)

        def _wait_fn(kind: str, timeout: float = 600) -> dict:
            job.awaiting = kind
            job.gate_payload = {}
            job.gate.clear()
            emit(job, "status", awaiting=kind)
            ok = job.gate.wait(timeout=timeout)
            job.awaiting = ""
            if not ok:
                return {"skip": True, "timeout": True}
            return dict(job.gate_payload or {})

        kind = getattr(job.parsed, "kind", "") or ""
        if kind == "extras":
            report = run_extras_pipeline(
                slug=job.parsed.slug or "chatguru",
                company=job.parsed.company or "",
                emit=_emit,
                set_step=_set_step,
                demo=bool(job.parsed.demo),
                preferred_competitors=job.parsed.competitors,
                run_id=job.job_id,
            )
        else:
            report = run_pipeline(
                job.parsed, _emit, _set_step, run_id=job.job_id, wait_fn=_wait_fn
            )
        job.report = report
        job.status = "done"
        try:
            usage_db.finish_run(
                job.job_id,
                status="done",
                report_summary=usage_db.summarize_report(report if isinstance(report, dict) else None),
            )
        except Exception:
            pass
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
            try:
                usage_db.upsert_step(
                    job.job_id, job.current_step, "error", str(e), error_message=str(e)
                )
            except Exception:
                pass
        try:
            usage_db.finish_run(job.job_id, status="error", error_message=str(e))
        except Exception:
            pass
        emit(job, "error", message=str(e), steps=job.steps)


@app.get("/")
def index():
    resp = send_from_directory(STATIC_DIR, "index.html")
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    return resp


@app.get("/video")
def video_demo():
    """Demo gravavel: Chatguru + canais extra, sem paywalls Pro."""
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
        examples = ["seusite.com.br"]
    elif live_ok:
        examples = ["seusite.com.br"]
    else:
        missing = ", ".join(_missing_live_keys())
        ask = (
            "Para iniciarmos, envie o endereço do seu site logo abaixo. "
            f"(Servidor ainda sem todas as API keys: {missing}.)"
        )
        examples = ["seusite.com.br"]
    return jsonify({
        "greeting": greeting,
        "ask": ask,
        "examples": examples,
        "steps": STEP_DEFS,
        "demo_only": DEMO_ONLY,
        "live_ready": live_ok,
        "keys_ready": keys,
        "payments_ready": mp_configured(),
        "meta_capi_ready": capi_configured(),
        "pro_price": PRO_PRICE,
        "extras_price": EXTRAS_PRICE,
        "version": APP_VERSION,
    })


@app.post("/api/meta/event")
def meta_capi_event():
    """Relay browser events to Meta Conversions API (bypass adblock/Firefox ETP)."""
    data = request.get_json(silent=True) or {}
    event_name = (data.get("event_name") or data.get("event") or "").strip()
    if not event_name:
        return jsonify({"ok": False, "error": "event_name obrigatorio"}), 400

    custom = data.get("custom_data")
    if custom is not None and not isinstance(custom, dict):
        custom = {}

    client_ip = (
        (request.headers.get("CF-Connecting-IP") or "").strip()
        or (request.headers.get("X-Forwarded-For") or "").split(",")[0].strip()
        or (request.remote_addr or "")
    )
    result = send_capi_event(
        event_name,
        event_source_url=str(data.get("event_source_url") or request.referrer or "")[:2048],
        custom_data=custom if isinstance(custom, dict) else {},
        client_ip=client_ip,
        user_agent=(request.headers.get("User-Agent") or "")[:512],
        event_id=str(data.get("event_id") or "").strip(),
        test_event_code=str(data.get("test_event_code") or "").strip(),
    )
    status = 200 if result.get("ok") or result.get("skipped") else 502
    return jsonify(result), status


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
    try:
        usage_db.create_run(
            run_id=job_id,
            kind="free",
            company=parsed.company or "",
            slug=parsed.slug or "",
            url=parsed.url or "",
            demo=bool(parsed.demo),
            preferred_competitors=parsed.competitors or [],
            user_agent=request.headers.get("User-Agent", "")[:300],
            ip=(request.headers.get("X-Forwarded-For") or request.remote_addr or "")[:80],
        )
        usage_db.append_log(job_id, f"Uso iniciado: {parsed.company or parsed.slug} ({parsed.url or 'sem url'})")
    except Exception as e:
        print(f"[usage_db] create_run falhou: {e}")

    threading.Thread(target=_worker, args=(job,), daemon=True).start()

    # Lead via Conversions API (servidor) — nao depende do Pixel no browser
    meta_lead = None
    try:
        meta_lead = track_lead_from_request(
            company=parsed.company or "",
            slug=parsed.slug or "",
            url=parsed.url or "",
            demo=bool(parsed.demo),
            request_obj=request,
            event_source_url=(
                os.environ.get("PUBLIC_BASE_URL", "").strip().rstrip("/") + "/"
            ),
        )
    except Exception as e:
        print(f"[meta_capi] Lead falhou: {e}", flush=True)
        meta_lead = {"ok": False, "error": str(e)}

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
        "meta_lead": meta_lead,
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
    """Volta do Checkout Pro: avisa a aba do chat e tenta fechar o popup."""
    status = (request.args.get("status") or "").strip()
    order_id = (request.args.get("order_id") or "").strip()
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

    if status == "success":
        title = "Pagamento aprovado"
        body = (
            "Tudo certo. Pode fechar esta janela — o chat do Radar já deve "
            "pedir seu e-mail ou WhatsApp na aba que ficou aberta."
        )
    elif status == "pending":
        title = "Pagamento pendente"
        body = "Se pagou via PIX, aguarde a confirmação e volte à aba do chat."
    else:
        title = "Pagamento não concluído"
        body = "Feche esta janela e tente de novo pelo chat, se quiser."

    # Escape for JS string literals
    safe_status = status.replace("\\", "\\\\").replace("'", "\\'")
    safe_order = order_id.replace("\\", "\\\\").replace("'", "\\'")

    html = f"""<!doctype html><html lang="pt-BR"><head>
<meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>{title}</title>
<style>
body{{font-family:system-ui,sans-serif;background:#fafafa;color:#111;display:flex;min-height:100vh;align-items:center;justify-content:center;margin:0}}
.card{{background:#fff;border:1px solid #e5e7eb;border-radius:16px;padding:28px 24px;max-width:440px;text-align:center;box-shadow:0 8px 30px rgba(0,0,0,.06)}}
h1{{font-size:18px;margin:0 0 8px}}p{{color:#555;font-size:14px;line-height:1.5;margin:0 0 8px}}
button{{display:inline-block;margin-top:12px;padding:10px 16px;border-radius:10px;border:0;background:#4f46e5;color:#fff;font-weight:700;font-size:14px;cursor:pointer}}
.hint{{font-size:12px;color:#8c8ca0;margin-top:14px}}
</style>
</head><body><div class="card">
<h1>{title}</h1>
<p>{body}</p>
<button type="button" id="backBtn">Voltar ao chat (fechar esta janela)</button>
<p class="hint" id="hint">Se o botão não fechar, volte manualmente à aba do Radar que já estava aberta — não abra o site de novo.</p>
</div>
<script>
(function () {{
  var status = '{safe_status}';
  var orderId = '{safe_order}';
  function notifyOpener() {{
    try {{
      if (window.opener && !window.opener.closed) {{
        window.opener.postMessage({{ type: 'mp_return', status: status, order_id: orderId }}, '*');
        try {{ window.opener.focus(); }} catch (e) {{}}
        return true;
      }}
    }} catch (e) {{}}
    return false;
  }}
  function goBack() {{
    notifyOpener();
    // Tenta fechar o popup do Mercado Pago
    setTimeout(function () {{
      try {{ window.close(); }} catch (e) {{}}
      document.getElementById('hint').textContent =
        'Não foi possível fechar automaticamente. Feche esta aba e continue no chat que já estava aberto.';
    }}, 120);
  }}
  document.getElementById('backBtn').addEventListener('click', goBack);
  // Auto: avisa o chat e tenta fechar em ~1s
  notifyOpener();
  setTimeout(goBack, 900);
}})();
</script>
</body></html>"""
    return Response(html, mimetype="text/html")


@app.post("/api/leads")
def create_lead():
    """Captura nome + e-mail/WhatsApp (antes ou depois do pagamento)."""
    data = request.get_json(silent=True) or {}
    contact_type = (data.get("contact_type") or "").strip().lower()
    contact = (data.get("contact") or "").strip()
    name = (data.get("name") or "").strip()
    channel = (data.get("channel") or "").strip()
    order_id = (data.get("order_id") or "").strip()
    pre_payment = bool(data.get("pre_payment"))
    lead_id = (data.get("lead_id") or "").strip() or uuid.uuid4().hex[:12]

    if not contact or len(contact) < 5:
        return jsonify({"error": "Informe um contato valido (e-mail ou WhatsApp)"}), 400

    if contact_type not in ("email", "whatsapp"):
        if "@" in contact:
            contact_type = "email"
        else:
            digits = "".join(ch for ch in contact if ch.isdigit())
            if len(digits) >= 10:
                contact_type = "whatsapp"
            else:
                return jsonify({
                    "error": "Envie um e-mail (ex: seu@email.com) ou WhatsApp com DDD (ex: 11999999999).",
                }), 400

    paid = False
    status = "awaiting_payment" if pre_payment else "pending_manual"

    if pre_payment:
        paid = False
        status = "awaiting_payment"
    elif mp_configured():
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
        status = "pending_manual"
        mark_order_contact(order_id, contact_type, contact)
    elif order_id:
        order = load_order(order_id)
        if order and order.get("status") == "paid":
            paid = True
            status = "pending_manual"
            mark_order_contact(order_id, contact_type, contact)

    lead = {
        "id": lead_id,
        "ts": time.time(),
        "offer": data.get("offer") or "deep_channel",
        "channel": channel,
        "name": name,
        "contact_type": contact_type,
        "contact": contact,
        "company": (data.get("company") or "").strip(),
        "slug": (data.get("slug") or "").strip(),
        "job_id": (data.get("job_id") or "").strip(),
        "order_id": order_id,
        "simulated_payment": not paid and not pre_payment,
        "status": status,
        "paid": paid,
        "pre_payment": pre_payment,
    }
    LEADS_DIR.mkdir(parents=True, exist_ok=True)
    with LEADS_FILE.open("a", encoding="utf-8") as f:
        f.write(json.dumps(lead, ensure_ascii=False) + "\n")
    try:
        usage_db.save_lead(lead)
    except Exception as e:
        print(f"[LEAD] sqlite save failed: {e}")
    print(
        f"[LEAD] {lead['id']} paid={paid} pre={pre_payment} "
        f"{lead['channel']} name={name!r} {lead['contact_type']}={lead['contact']}"
    )
    label = "e-mail" if contact_type == "email" else "WhatsApp"
    if pre_payment:
        msg = (
            f"Dados salvos ({name or 'lead'} / {label}). "
            "Segue o link do Mercado Pago para concluir o pagamento."
        )
    else:
        msg = (
            f"Perfeito. Registrei seu {label} (**{contact}**) para a análise Pro de "
            f"**{channel or 'canal'}**. Em até 24h você recebe o relatório."
        )
    return jsonify({
        "ok": True,
        "lead_id": lead["id"],
        "paid": paid,
        "name": name,
        "contact_type": contact_type,
        "status": status,
        "message": msg,
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
    try:
        usage_db.create_run(
            run_id=job_id,
            kind="extras",
            company=company,
            slug=slug,
            url="",
            demo=demo,
            parent_run_id=job_id_src,
            user_agent=request.headers.get("User-Agent", "")[:300],
            ip=(request.headers.get("X-Forwarded-For") or request.remote_addr or "")[:80],
        )
        usage_db.append_log(job_id, f"Canais extra iniciados: {company} / {slug}")
    except Exception as e:
        print(f"[usage_db] create_run extras falhou: {e}")
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
            f"Vou analisar os **canais extra** de **{company}**: Meta Ads, Instagram, YouTube e TikTok "
            "(LinkedIn em breve). Resultados aparecem etapa a etapa."
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
        # Se o job ja terminou (ex.: enrich rapido / SSE atrasado), reenvia o resultado
        if job.status == "done":
            yield _sse({
                "type": "complete",
                "job_id": job_id,
                "report": job.report,
                "enrich": (getattr(job.parsed, "kind", "") == "enrich"),
                "progressive": True,
            })
            return
        if job.status == "error":
            yield _sse({"type": "error", "job_id": job_id, "message": job.error or "erro"})
            return
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


@app.post("/api/competitors/enrich")
def enrich_competitors():
    """Adiciona ate 3 concorrentes e reprocessa Ads/SEO/Marca sem travar o fluxo principal."""
    data = request.get_json(silent=True) or {}
    parent_job_id = (data.get("job_id") or "").strip()
    parent = jobs.get(parent_job_id) if parent_job_id else None

    slug = (data.get("slug") or "").strip()
    company = (data.get("company") or "").strip()
    if parent and parent.parsed:
        slug = slug or (parent.parsed.slug or "")
        company = company or (parent.parsed.company or "")
    if not slug and company:
        from parse_input import slugify_client
        slug = slugify_client(company)
    slug = slug or "cliente"
    company = company or slug
    demo = bool(data.get("demo"))
    raw = data.get("competitors")
    comps: list[str] = []
    if isinstance(raw, list):
        comps = [str(x).strip() for x in raw if str(x).strip()]
    elif isinstance(raw, str):
        from parse_input import _split_competitors
        comps = _split_competitors(raw)
    comps = comps[:3]
    if not comps:
        return jsonify({"ok": True, "skipped": True})

    preferred = list(data.get("preferred") or [])
    if parent and parent.parsed and parent.parsed.competitors:
        for c in parent.parsed.competitors:
            if c not in preferred:
                preferred.append(c)

    job_id = uuid.uuid4().hex[:12]
    parsed = ParsedInput(
        company=company,
        slug=slug,
        demo=demo or DEMO_ONLY,
        competitors=preferred,
        kind="enrich",
        raw=f"enrich:{slug}",
    )
    job = Job(job_id=job_id, parsed=parsed)
    jobs[job_id] = job

    def _worker_enrich() -> None:
        try:
            job.status = "running"

            def _emit(event_type: str, **payload: Any) -> None:
                emit(job, event_type, **payload)
                # Espelha partials no job pai se ainda existir (mesma aba SSE)
                if parent and event_type in ("partial", "log", "progress"):
                    try:
                        emit(parent, event_type, **payload)
                    except Exception:
                        pass

            result = run_enrich_competitors(
                slug=slug,
                company=company,
                competitors=comps,
                emit=_emit,
                preferred_competitors=preferred,
                demo=bool(parsed.demo),
            )
            if parent and parent.parsed:
                for n in comps:
                    if n not in parent.parsed.competitors:
                        parent.parsed.competitors.append(n)
                if isinstance(result, dict) and result.get("report"):
                    parent.report = result["report"]
            job.report = (result or {}).get("report")
            job.status = "done"
            emit(job, "enrich_complete", added=comps, report=job.report)
            emit(job, "complete", report=job.report, enrich=True)
            if parent:
                emit(parent, "enrich_complete", added=comps, report=job.report)
        except Exception as e:
            job.status = "error"
            job.error = str(e)
            emit(job, "error", message=str(e))
            if parent:
                emit(parent, "log", line=f"Falha ao enriquecer concorrentes: {e}")
                emit(parent, "enrich_error", message=str(e))

    try:
        usage_db.create_run(
            run_id=job_id,
            kind="enrich",
            company=company,
            slug=slug,
            demo=bool(parsed.demo),
            preferred_competitors=comps,
            parent_run_id=parent_job_id,
        )
    except Exception:
        pass
    threading.Thread(target=_worker_enrich, daemon=True).start()
    return jsonify({
        "ok": True,
        "job_id": job_id,
        "parent_job_id": parent_job_id or None,
        "competitors": comps,
        "ack": f"Vou incluir **{', '.join(comps)}** e atualizar Ads, SEO e Marca.",
    })


@app.post("/api/jobs/<job_id>/continue")
def continue_job(job_id: str):
    """Retoma pipeline pausada (ex.: adicionar concorrentes apos briefing)."""
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job nao encontrado"}), 404
    data = request.get_json(silent=True) or {}
    kind = (data.get("kind") or job.awaiting or "").strip()
    if job.awaiting and kind and kind != job.awaiting:
        return jsonify({"error": f"Job aguarda '{job.awaiting}', nao '{kind}'"}), 409
    if not job.awaiting:
        # Idempotente: ja retomou
        return jsonify({"ok": True, "already": True})

    if kind == "extra_competitors":
        raw = data.get("competitors")
        comps: list[str] = []
        if isinstance(raw, list):
            comps = [str(x).strip() for x in raw if str(x).strip()]
        elif isinstance(raw, str):
            from parse_input import _split_competitors
            comps = _split_competitors(raw)
        comps = comps[:3]
        skip = bool(data.get("skip")) or not comps
        job.gate_payload = {"competitors": comps, "skip": skip}
    else:
        job.gate_payload = dict(data)
    job.gate.set()
    return jsonify({"ok": True, "kind": kind, "payload": job.gate_payload})


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


def _admin_authorized() -> bool:
    expected = (os.environ.get("ADMIN_DASHBOARD_TOKEN") or "").strip()
    if not expected:
        # Sem token: libera so em localhost (dev)
        return (request.remote_addr or "") in ("127.0.0.1", "::1")
    token = (
        request.args.get("token")
        or request.headers.get("X-Admin-Token")
        or request.cookies.get("radar_admin_token")
        or ""
    ).strip()
    return token == expected


@app.get("/admin")
def admin_dashboard():
    if not _admin_authorized():
        return (
            "<h1>403</h1><p>Defina ADMIN_DASHBOARD_TOKEN no Render e acesse "
            "<code>/admin?token=SEU_TOKEN</code>.</p>",
            403,
        )
    resp = send_from_directory(STATIC_DIR, "admin.html")
    resp.headers["Cache-Control"] = "no-store"
    token = (request.args.get("token") or "").strip()
    if token:
        resp.set_cookie("radar_admin_token", token, httponly=True, samesite="Lax")
    return resp


@app.get("/api/admin/runs")
def admin_list_runs():
    if not _admin_authorized():
        return jsonify({"error": "unauthorized"}), 403
    limit = min(int(request.args.get("limit") or 50), 200)
    offset = max(int(request.args.get("offset") or 0), 0)
    return jsonify({"runs": usage_db.list_runs(limit=limit, offset=offset)})


@app.get("/api/admin/runs/<run_id>")
def admin_get_run(run_id: str):
    if not _admin_authorized():
        return jsonify({"error": "unauthorized"}), 403
    run = usage_db.get_run(run_id)
    if not run:
        return jsonify({"error": "Run nao encontrado"}), 404
    return jsonify({
        "run": run,
        "steps": usage_db.get_run_steps(run_id),
        "artifacts": usage_db.get_run_artifacts(run_id),
        "logs": usage_db.get_run_logs(run_id, limit=800),
    })


@app.get("/api/admin/runs/<run_id>/file")
def admin_get_file(run_id: str):
    """Serve artifact file by id or relative path under data/runs/<id>/."""
    if not _admin_authorized():
        return jsonify({"error": "unauthorized"}), 403
    art_id = request.args.get("artifact_id")
    rel = (request.args.get("path") or "").strip()
    path: Optional[Path] = None
    if art_id:
        for a in usage_db.get_run_artifacts(run_id):
            if str(a.get("id")) == str(art_id):
                path = Path(a["path"])
                break
    elif rel:
        # only allow under runs/<run_id>
        candidate = (usage_db.RUNS_DIR / run_id / rel).resolve()
        root = (usage_db.RUNS_DIR / run_id).resolve()
        if str(candidate).startswith(str(root)) and candidate.exists():
            path = candidate
    if not path or not path.exists():
        return jsonify({"error": "Arquivo nao encontrado"}), 404
    return send_from_directory(str(path.parent), path.name, as_attachment=True)


if __name__ == "__main__":
    usage_db.init_db()
    port = int(os.environ.get("PORT", "8766"))
    print(f"Radar da Concorrencia -> http://127.0.0.1:{port}  DEMO_ONLY={DEMO_ONLY} live_keys={_live_keys_ready()}")
    print(f"Admin dashboard -> http://127.0.0.1:{port}/admin")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
