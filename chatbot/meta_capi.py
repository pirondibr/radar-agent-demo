# -*- coding: utf-8 -*-
"""Meta Conversions API (CAPI) — eventos pelo servidor (sem depender do browser).

Docs: https://developers.facebook.com/docs/marketing-api/conversions-api/using-the-api/
"""

from __future__ import annotations

import os
import time
import uuid
from typing import Any, Optional

import requests

GRAPH_VERSION = os.environ.get("META_GRAPH_VERSION", "v21.0").strip() or "v21.0"
DEFAULT_PIXEL_ID = "2289325455219613"


def meta_pixel_id() -> str:
    return os.environ.get("META_PIXEL_ID", "").strip() or DEFAULT_PIXEL_ID


def capi_access_token() -> str:
    return (
        os.environ.get("META_CAPI_ACCESS_TOKEN", "").strip()
        or os.environ.get("META_PIXEL_ACCESS_TOKEN", "").strip()
    )


def capi_configured() -> bool:
    return bool(capi_access_token() and meta_pixel_id())


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
    pixel_id = meta_pixel_id()
    if not token:
        return {
            "ok": False,
            "skipped": True,
            "reason": "META_CAPI_ACCESS_TOKEN ausente",
            "pixel_id": pixel_id,
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

    url = f"https://graph.facebook.com/{GRAPH_VERSION}/{pixel_id}/events"
    try:
        resp = requests.post(
            url,
            params={"access_token": token},
            json=body,
            timeout=20,
        )
    except Exception as e:
        print(f"[meta_capi] erro de rede: {e}", flush=True)
        return {"ok": False, "error": str(e), "event_id": event_id, "pixel_id": pixel_id}

    try:
        data = resp.json()
    except Exception:
        data = {"raw": (resp.text or "")[:500]}

    ok = resp.status_code < 400 and int(data.get("events_received") or 0) >= 1
    print(
        f"[meta_capi] event={event_name} ok={ok} status={resp.status_code} "
        f"test={bool(test_code)} resp={str(data)[:200]}",
        flush=True,
    )
    return {
        "ok": ok,
        "status_code": resp.status_code,
        "event_id": event_id,
        "pixel_id": pixel_id,
        "test_event_code": test_code or None,
        "response": data,
    }


def track_lead_from_request(
    *,
    company: str = "",
    slug: str = "",
    url: str = "",
    demo: bool = False,
    request_obj: Any = None,
    event_source_url: str = "",
) -> dict[str, Any]:
    """Lead padrao: usuario pediu analise de um site."""
    client_ip = ""
    user_agent = ""
    if request_obj is not None:
        client_ip = (
            (request_obj.headers.get("CF-Connecting-IP") or "").strip()
            or (request_obj.headers.get("X-Forwarded-For") or "").split(",")[0].strip()
            or (request_obj.remote_addr or "")
        )
        user_agent = (request_obj.headers.get("User-Agent") or "")[:512]
        if not event_source_url:
            event_source_url = (
                (request_obj.headers.get("Origin") or "").strip()
                or (request_obj.referrer or "")
                or os.environ.get("PUBLIC_BASE_URL", "").strip()
            )

    return send_capi_event(
        "Lead",
        event_source_url=(event_source_url or "")[:2048],
        custom_data={
            "content_name": "site_analysis_request",
            "content_category": "radar_analysis",
            "status": "demo" if demo else "live",
            "content_ids": [slug or company or "site"][:1],
        },
        client_ip=client_ip,
        user_agent=user_agent,
    )
