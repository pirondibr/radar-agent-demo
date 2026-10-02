# -*- coding: utf-8 -*-
"""Paddle Billing checkout (Winner Tracker EN) — overlay + webhooks."""

from __future__ import annotations

import hashlib
import hmac
import json
import os
import time
import uuid
from typing import Any, Optional

from mercadopago_client import load_order, save_order

# Sandbox catalog created via Paddle MCP (override with env in live)
DEFAULT_PRICE_PRO = "pri_01m3ywxjq2sevxc9ydrp0ad5gn"
DEFAULT_PRICE_EXTRAS = "pri_01m3ywxk2v8h34yz38k1mbg8m9"
DEFAULT_CLIENT_TOKEN = "test_d7daca219bacd1660300a0318b0"

PRO_PRICE_USD = float(os.environ.get("PRO_PRICE_USD", "29").strip() or "29")
EXTRAS_PRICE_USD = float(os.environ.get("EXTRAS_PRICE_USD", "19").strip() or "19")


def paddle_api_key() -> str:
    return (
        os.environ.get("PADDLE_API_KEY", "").strip()
        or os.environ.get("PADDLE_SANDBOX_API_KEY", "").strip()
    )


def paddle_client_token() -> str:
    return (
        os.environ.get("PADDLE_CLIENT_TOKEN", "").strip()
        or os.environ.get("PADDLE_CLIENT_SIDE_TOKEN", "").strip()
        or DEFAULT_CLIENT_TOKEN
    )


def paddle_webhook_secret() -> str:
    return (
        os.environ.get("PADDLE_WEBHOOK_SECRET", "").strip()
        or os.environ.get("PADDLE_ENDPOINT_SECRET_KEY", "").strip()
    )


def paddle_sandbox() -> bool:
    flag = os.environ.get("PADDLE_ENV", os.environ.get("PADDLE_SANDBOX", "sandbox")).strip().lower()
    if flag in ("live", "production", "0", "false", "no"):
        return False
    token = paddle_client_token()
    if token.startswith("live_"):
        return False
    return True


def paddle_configured() -> bool:
    return bool(paddle_client_token())


def price_id_for_product(product: str) -> str:
    product = (product or "deep_channel").strip().lower()
    if product == "extras_pack":
        return (
            os.environ.get("PADDLE_PRICE_EXTRAS", "").strip()
            or DEFAULT_PRICE_EXTRAS
        )
    return (
        os.environ.get("PADDLE_PRICE_PRO", "").strip()
        or DEFAULT_PRICE_PRO
    )


def amount_for_product(product: str) -> float:
    product = (product or "deep_channel").strip().lower()
    if product == "extras_pack":
        return EXTRAS_PRICE_USD
    return PRO_PRICE_USD


def create_paddle_checkout(
    *,
    channel: str,
    company: str = "",
    slug: str = "",
    job_id: str = "",
    product: str = "deep_channel",
    funnel: str = "",
    customer_email: str = "",
) -> dict[str, Any]:
    """Cria ordem local e devolve dados para abrir Paddle.js overlay."""
    if not paddle_configured():
        raise RuntimeError("PADDLE_CLIENT_TOKEN ausente")

    product = (product or "deep_channel").strip().lower()
    price_id = price_id_for_product(product)
    amount = amount_for_product(product)
    order_id = uuid.uuid4().hex[:12]
    order = {
        "id": order_id,
        "provider": "paddle",
        "status": "pending",
        "product": product,
        "channel": channel,
        "company": company,
        "slug": slug,
        "job_id": job_id,
        "funnel": funnel,
        "amount": amount,
        "currency": "USD",
        "price_id": price_id,
        "customer_email": (customer_email or "").strip(),
        "created_at": time.time(),
        "sandbox": paddle_sandbox(),
    }
    save_order(order)
    return {
        "provider": "paddle",
        "order_id": order_id,
        "price_id": price_id,
        "client_token": paddle_client_token(),
        "sandbox": paddle_sandbox(),
        "amount": amount,
        "currency": "USD",
        "channel": channel,
        "product": product,
        # Paddle custom_data: flat string values only; omit empties
        "custom_data": {
            k: str(v)
            for k, v in {
                "order_id": order_id,
                "product": product,
                "channel": channel,
                "company": company,
                "slug": slug,
                "job_id": job_id,
                "funnel": funnel,
            }.items()
            if str(v or "").strip()
        },
    }


def verify_paddle_signature(raw_body: bytes, signature_header: str, secret: str) -> bool:
    """Valida header Paddle-Signature (ts=...;h1=...)."""
    if not secret or not signature_header:
        return False
    parts: dict[str, str] = {}
    for chunk in signature_header.split(";"):
        chunk = chunk.strip()
        if "=" not in chunk:
            continue
        k, _, v = chunk.partition("=")
        parts[k.strip()] = v.strip()
    ts = parts.get("ts") or ""
    h1 = parts.get("h1") or ""
    if not ts or not h1:
        return False
    try:
        if abs(time.time() - int(ts)) > 300:
            return False
    except Exception:
        return False
    payload = f"{ts}:{raw_body.decode('utf-8')}".encode("utf-8")
    digest = hmac.new(secret.encode("utf-8"), payload, hashlib.sha256).hexdigest()
    return hmac.compare_digest(digest, h1)


def apply_paddle_transaction(event: dict[str, Any]) -> Optional[dict[str, Any]]:
    """Marca ordem local como paga a partir de event transaction.*."""
    data = event.get("data") if isinstance(event.get("data"), dict) else event
    if not isinstance(data, dict):
        return None
    custom = data.get("custom_data") if isinstance(data.get("custom_data"), dict) else {}
    order_id = str(custom.get("order_id") or "").strip()
    if not order_id:
        # fallback: custom_data may be nested under details
        details = data.get("details") if isinstance(data.get("details"), dict) else {}
        custom2 = details.get("custom_data") if isinstance(details.get("custom_data"), dict) else {}
        order_id = str(custom2.get("order_id") or "").strip()
    if not order_id:
        return None
    order = load_order(order_id)
    if not order:
        return None

    status = str(data.get("status") or "").lower()
    event_type = str(event.get("event_type") or event.get("type") or "").lower()
    order["paddle_status"] = status
    order["paddle_transaction_id"] = str(data.get("id") or order.get("paddle_transaction_id") or "")
    newly_paid = False
    paid_like = status in ("completed", "paid") or "completed" in event_type or event_type.endswith(".paid")
    if paid_like:
        newly_paid = order.get("status") != "paid"
        order["status"] = "paid"
        order["paid_at"] = time.time()
        try:
            totals = data.get("details", {}).get("totals") if isinstance(data.get("details"), dict) else None
            if isinstance(totals, dict) and totals.get("grand_total") is not None:
                # Paddle amounts are in lowest denomination as string
                cents = int(str(totals.get("grand_total")))
                order["amount"] = cents / 100.0
                order["currency"] = str(totals.get("currency_code") or order.get("currency") or "USD")
        except Exception:
            pass
    save_order(order)

    if newly_paid and not order.get("meta_purchase_sent"):
        try:
            from meta_capi import track_purchase_from_order

            result = track_purchase_from_order(order)
            order["meta_purchase_sent"] = bool(result.get("ok") or result.get("skipped"))
            order["meta_purchase"] = result
            save_order(order)
        except Exception as e:
            print(f"[paddle] meta purchase failed order={order_id}: {e}", flush=True)

    if newly_paid and not order.get("openai_purchase_sent"):
        try:
            from openai_capi import send_event as send_openai_event

            sale_data: dict[str, Any] = {"type": "custom"}
            try:
                sale_data["amount"] = int(round(float(order.get("amount") or 0) * 100))
                sale_data["currency"] = str(order.get("currency") or "USD")
            except Exception:
                pass
            result = send_openai_event(
                "custom",
                data=sale_data,
                custom_event_name="sale",
            )
            order["openai_purchase_sent"] = bool(result.get("ok") or result.get("skipped"))
            order["openai_purchase"] = result
            save_order(order)
        except Exception as e:
            print(f"[paddle] openai sale failed order={order_id}: {e}", flush=True)

    return order


def public_config() -> dict[str, Any]:
    return {
        "provider": "paddle",
        "configured": paddle_configured(),
        "sandbox": paddle_sandbox(),
        "client_token": paddle_client_token() if paddle_configured() else "",
        "price_pro": price_id_for_product("deep_channel"),
        "price_extras": price_id_for_product("extras_pack"),
        "pro_price": PRO_PRICE_USD,
        "extras_price": EXTRAS_PRICE_USD,
        "currency": "USD",
    }
