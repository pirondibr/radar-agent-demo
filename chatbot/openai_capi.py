# -*- coding: utf-8 -*-
"""OpenAI Ads Conversions API — eventos pelo servidor.

Docs: https://developers.openai.com/ads/conversions-api
"""

from __future__ import annotations

import hashlib
import os
import time
from typing import Any, Optional

import requests

DEFAULT_PIXEL_ID = "LpS3eYX3NmAuftchn5HVwA"
EVENTS_URL = "https://bzr.openai.com/v1/events"
ALLOWED_EVENTS = {
    "page_viewed",
    "lead_created",
    "checkout_started",
    "order_created",
}


def pixel_id() -> str:
    return os.environ.get("OPENAI_PIXEL_ID", "").strip() or DEFAULT_PIXEL_ID


def api_key() -> str:
    return os.environ.get("OPENAI_CONVERSIONS_API_KEY", "").strip()


def configured() -> bool:
    return bool(api_key() and pixel_id())


def _sha256(value: str) -> str:
    text = (value or "").strip().lower()
    if not text:
        return ""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _phone_digits(phone: str) -> str:
    digits = "".join(ch for ch in (phone or "") if ch.isdigit())
    digits = digits.lstrip("0")
    if len(digits) >= 10 and not digits.startswith("55"):
        digits = "55" + digits
    if len(digits) < 8 or len(digits) > 15:
        return ""
    return digits


def send_event(
    event_type: str,
    *,
    event_id: str = "",
    source_url: str = "",
    data: Optional[dict[str, Any]] = None,
    email: str = "",
    phone: str = "",
    client_ip: str = "",
    user_agent: str = "",
) -> dict[str, Any]:
    """Envia um evento. A chave fica só no header, nunca no retorno."""
    kind = (event_type or "").strip()
    if kind not in ALLOWED_EVENTS:
        return {"ok": False, "error": "event_type invalido"}
    key = api_key()
    if not key:
        return {"ok": False, "skipped": True, "reason": "not_configured"}

    payload = dict(data or {})
    if kind == "lead_created":
        payload["type"] = "customer_action"
        payload.pop("contents", None)
    else:
        payload.setdefault("type", "contents")
    if payload.get("amount") is not None:
        try:
            payload["amount"] = int(payload["amount"])
        except (TypeError, ValueError):
            payload.pop("amount", None)
        else:
            payload.setdefault("currency", "BRL")

    event_key = (event_id or f"{kind}_{int(time.time() * 1000)}")[:128]
    event: dict[str, Any] = {
        "id": event_key,
        "type": kind,
        "timestamp_ms": int(time.time() * 1000),
        "action_source": "web",
        "source_url": (source_url or "https://agente.radardaconcorrencia.com.br/")[:2048],
        "data": payload,
    }
    user: dict[str, Any] = {}
    email_hash = _sha256(email)
    if email_hash:
        user["emails_sha256"] = [email_hash]
    phone_hash = _sha256(_phone_digits(phone))
    if phone_hash:
        user["phone_numbers_sha256"] = [phone_hash]
    if client_ip:
        user["ip_address"] = client_ip[:64]
    if user_agent:
        user["user_agent"] = user_agent[:512]
    if user:
        event["user"] = user

    try:
        resp = requests.post(
            EVENTS_URL,
            params={"pid": pixel_id()},
            headers={
                "Authorization": f"Bearer {key}",
                "Content-Type": "application/json",
                "User-Agent": "radar-agent/1.0",
            },
            json={"validate_only": False, "events": [event]},
            timeout=8,
        )
    except requests.RequestException:
        return {"ok": False, "error": "request_failed", "event_id": event_key}

    body: dict[str, Any] = {}
    try:
        parsed = resp.json()
        if isinstance(parsed, dict):
            body = parsed
    except ValueError:
        body = {}
    ok = resp.status_code < 300
    return {
        "ok": ok,
        "status": resp.status_code,
        "event_id": event_key,
        "accepted": body.get("accepted_events"),
        "error": None if ok else "rejected",
    }
