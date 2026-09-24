# -*- coding: utf-8 -*-
"""Mercado Pago Checkout Pro (PIX + cartão) para Radar Pro."""

from __future__ import annotations

import json
import os
import time
import uuid
from pathlib import Path
from typing import Any, Optional

import requests

MP_API = "https://api.mercadopago.com"
PRO_PRICE = float(os.environ.get("PRO_PRICE_BRL", "99").strip() or "99")
PRO_TITLE = os.environ.get("PRO_TITLE", "Radar Pro - Analise profunda de canal")
EXTRAS_TITLE = os.environ.get(
    "EXTRAS_TITLE",
    "Radar - Canais extra (Meta, LinkedIn, Instagram, YouTube)",
)
EXTRAS_PRICE = float(
    os.environ.get("EXTRAS_PRICE_BRL", os.environ.get("PRO_PRICE_BRL", "99")).strip() or "99"
)

ORDERS_DIR = Path(__file__).resolve().parent / "data" / "orders"
ORDERS_DIR.mkdir(parents=True, exist_ok=True)


def mp_access_token() -> str:
    return (
        os.environ.get("MERCADOPAGO_ACCESS_TOKEN", "").strip()
        or os.environ.get("MP_ACCESS_TOKEN", "").strip()
    )


def mp_public_key() -> str:
    return (
        os.environ.get("MERCADOPAGO_PUBLIC_KEY", "").strip()
        or os.environ.get("MP_PUBLIC_KEY", "").strip()
    )


def mp_sandbox_mode() -> bool:
    """True when we should treat credentials as test.

    Note: modern MP test credentials use APP_USR-... (not only TEST-...).
    For APP_USR test tokens, Checkout Pro must open init_point (not
    sandbox_init_point) — sandbox_init_point often yields "Ops, ocorreu um erro".
    """
    flag = os.environ.get("MERCADOPAGO_SANDBOX", "").strip().lower()
    if flag in ("1", "true", "yes"):
        return True
    if flag in ("0", "false", "no"):
        return False
    token = mp_access_token().upper()
    return token.startswith("TEST-") or token.startswith("APP_USR-")


def mp_use_sandbox_init_point() -> bool:
    """Only classic TEST- tokens should open sandbox_init_point."""
    token = mp_access_token().upper()
    if token.startswith("TEST-"):
        return True
    # APP_USR test credentials (painel "Credenciais de teste") → init_point
    return False


def mp_configured() -> bool:
    return bool(mp_access_token())


def public_base_url() -> str:
    return (
        os.environ.get("PUBLIC_BASE_URL", "").strip().rstrip("/")
        or os.environ.get("RENDER_EXTERNAL_URL", "").strip().rstrip("/")
        or "http://127.0.0.1:8766"
    )


def _order_path(order_id: str) -> Path:
    return ORDERS_DIR / f"{order_id}.json"


def save_order(order: dict[str, Any]) -> None:
    ORDERS_DIR.mkdir(parents=True, exist_ok=True)
    _order_path(order["id"]).write_text(
        json.dumps(order, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_order(order_id: str) -> Optional[dict[str, Any]]:
    path = _order_path(order_id)
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def create_pro_checkout(
    *,
    channel: str,
    company: str = "",
    slug: str = "",
    job_id: str = "",
    product: str = "deep_channel",
) -> dict[str, Any]:
    """Cria preference Checkout Pro com PIX + cartão.

    product:
      - deep_channel: analise Pro de um canal (entrega manual 24h)
      - extras_pack: desbloqueia Meta/LinkedIn/IG/YouTube no chat
    """
    token = mp_access_token()
    if not token:
        raise RuntimeError(
            "MERCADOPAGO_ACCESS_TOKEN nao configurado. "
            "Defina no Render Environment (ou use o MCP get_credentials)."
        )

    order_id = uuid.uuid4().hex[:16]
    base = public_base_url()
    product = (product or "deep_channel").strip().lower()
    if product == "extras_pack":
        channel_label = "canais-extra"
        title = EXTRAS_TITLE
        price = EXTRAS_PRICE
        description = (
            f"Canais extra para {company or slug or 'cliente'}: "
            "Meta Ads, LinkedIn, Instagram e YouTube"
        )
    else:
        product = "deep_channel"
        channel_label = channel or "canal"
        title = f"{PRO_TITLE} ({channel_label})"
        price = PRO_PRICE
        description = (
            f"Analise Pro do canal {channel_label} "
            f"para {company or slug or 'cliente'}"
        )

    preference_body: dict[str, Any] = {
        "items": [
            {
                "id": f"radar-{product}-{channel_label}",
                "title": title[:120],
                "description": description[:256],
                "quantity": 1,
                "currency_id": "BRL",
                "unit_price": price,
            }
        ],
        "external_reference": order_id,
        "statement_descriptor": "RADAR PRO",
        "binary_mode": True,
        "payment_methods": {
            "excluded_payment_types": [
                {"id": "ticket"},
                {"id": "atm"},
            ],
            "installments": 12,
        },
        "back_urls": {
            "success": f"{base}/pay/return?status=success&order_id={order_id}",
            "failure": f"{base}/pay/return?status=failure&order_id={order_id}",
            "pending": f"{base}/pay/return?status=pending&order_id={order_id}",
        },
        "auto_return": "approved",
        "notification_url": f"{base}/api/webhooks/mercadopago?source_news=webhooks",
        "metadata": {
            "order_id": order_id,
            "channel": channel_label if product == "extras_pack" else channel,
            "company": company,
            "slug": slug,
            "job_id": job_id,
            "product": f"radar_{product}",
        },
    }

    resp = requests.post(
        f"{MP_API}/checkout/preferences",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=preference_body,
        timeout=45,
    )
    if resp.status_code >= 400:
        raise RuntimeError(f"Mercado Pago preference error {resp.status_code}: {resp.text[:500]}")

    pref = resp.json()
    use_sandbox_url = mp_use_sandbox_init_point()
    init_point = (
        (pref.get("sandbox_init_point") or pref.get("init_point"))
        if use_sandbox_url
        else (pref.get("init_point") or pref.get("sandbox_init_point"))
    )
    order = {
        "id": order_id,
        "created_at": time.time(),
        "status": "pending",
        "mp_status": "",
        "channel": channel_label if product == "extras_pack" else channel,
        "company": company,
        "slug": slug,
        "job_id": job_id,
        "amount": price,
        "currency": "BRL",
        "preference_id": pref.get("id"),
        "init_point": init_point,
        "sandbox_init_point": pref.get("sandbox_init_point"),
        "payment_id": "",
        "paid_at": None,
        "contact_type": "",
        "contact": "",
        "sandbox": mp_sandbox_mode(),
        "use_sandbox_url": use_sandbox_url,
        "product": product,
    }
    save_order(order)
    return {
        "order_id": order_id,
        "preference_id": order["preference_id"],
        "init_point": order["init_point"],
        "sandbox_init_point": order["sandbox_init_point"],
        "public_key": mp_public_key(),
        "amount": price,
        "channel": order["channel"],
        "product": product,
    }


def fetch_payment(payment_id: str) -> dict[str, Any]:
    token = mp_access_token()
    if not token:
        raise RuntimeError("MERCADOPAGO_ACCESS_TOKEN ausente")
    resp = requests.get(
        f"{MP_API}/v1/payments/{payment_id}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def apply_payment_to_order(payment: dict[str, Any]) -> Optional[dict[str, Any]]:
    """Atualiza ordem local a partir de um payment MP."""
    ext = str(payment.get("external_reference") or "").strip()
    meta = payment.get("metadata") or {}
    order_id = ext or str(meta.get("order_id") or "").strip()
    if not order_id:
        return None
    order = load_order(order_id)
    if not order:
        return None

    status = str(payment.get("status") or "").lower()
    order["mp_status"] = status
    order["payment_id"] = str(payment.get("id") or order.get("payment_id") or "")
    if status == "approved":
        order["status"] = "paid"
        order["paid_at"] = time.time()
    elif status in ("pending", "in_process", "in_mediation"):
        order["status"] = "pending"
    elif status in ("rejected", "cancelled", "refunded", "charged_back"):
        order["status"] = "failed"
    save_order(order)
    return order


def mark_order_contact(order_id: str, contact_type: str, contact: str) -> Optional[dict[str, Any]]:
    order = load_order(order_id)
    if not order:
        return None
    order["contact_type"] = contact_type
    order["contact"] = contact
    order["contact_at"] = time.time()
    save_order(order)
    return order
