# -*- coding: utf-8 -*-
"""Meta Conversions API (CAPI) — envia eventos pelo servidor.

Util quando o Pixel do browser e bloqueado (Firefox ETP, uBlock, etc.).
Docs: https://developers.facebook.com/docs/marketing-api/conversions-api/using-the-api/
"""

from __future__ import annotations

import os
import time
import uuid
from typing import Any, Optional

import requests

META_PIXEL_ID = (
    os.environ.get("META_PIXEL_ID", "").strip() or "2289325455219613"
)
GRAPH_VERSION = os.environ.get("META_GRAPH_VERSION", "v21.0").strip() or "v21.0"


def capi_access_token() -> str:
    return (
        os.environ.get("META_CAPI_ACCESS_TOKEN", "").strip()
        or os.environ.get("META_PIXEL_ACCESS_TOKEN", "").strip()
    )


def capi_configured() -> bool:
    return bool(capi_access_token() and META_PIXEL_ID)


def send_capi_event(
    event_name: str,
    *,
    event_source_url: str = "",
    custom_data: Optional[dict[str, Any]] = None,
    client_ip: str = "",
    user_agent: str = "",
    event_id: str = "",
    test_event_code: str = "",
) -> dict[str, Any]:
    token = capi_access_token()
    if not token:
        return {
            "ok": False,
            "skipped": True,
            "reason": "META_CAPI_ACCESS_TOKEN ausente",
        }

    event_id = (event_id or "").strip() or uuid.uuid4().hex
    test_code = (test_event_code or "").strip() or os.environ.get(
        "META_TEST_EVENT_CODE", ""
    ).strip()

    user_data: dict[str, Any] = {}
    if client_ip:
        user_data["client_ip_address"] = client_ip
    if user_agent:
        user_data["client_user_agent"] = user_agent

    event: dict[str, Any] = {
        "event_name": event_name,
        "event_time": int(time.time()),
        "action_source": "website",
        "event_id": event_id,
    }
    if event_source_url:
        event["event_source_url"] = event_source_url
    if user_data:
        event["user_data"] = user_data
    if custom_data:
        event["custom_data"] = custom_data

    body: dict[str, Any] = {"data": [event]}
    if test_code:
        body["test_event_code"] = test_code

    url = f"https://graph.facebook.com/{GRAPH_VERSION}/{META_PIXEL_ID}/events"
    try:
        resp = requests.post(
            url,
            params={"access_token": token},
            json=body,
            timeout=20,
        )
    except Exception as e:
        return {"ok": False, "error": str(e), "event_id": event_id}

    try:
        data = resp.json()
    except Exception:
        data = {"raw": (resp.text or "")[:500]}

    ok = resp.status_code < 400 and int(data.get("events_received") or 0) >= 1
    return {
        "ok": ok,
        "status_code": resp.status_code,
        "event_id": event_id,
        "pixel_id": META_PIXEL_ID,
        "test_event_code": test_code or None,
        "response": data,
    }
