# -*- coding: utf-8 -*-
"""Classifica origem do trafego (Google Ads, Meta, Organico, Direto, Outro)."""

from __future__ import annotations

from typing import Any, Optional
from urllib.parse import urlparse


def normalize_attribution(raw: Optional[dict[str, Any]]) -> dict[str, Any]:
    """Limpa e limita campos de atribuicao vindos do browser."""
    if not isinstance(raw, dict):
        return {}
    out: dict[str, Any] = {}
    for key in (
        "utm_source",
        "utm_medium",
        "utm_campaign",
        "utm_content",
        "utm_term",
        "gclid",
        "fbclid",
        "msclkid",
        "ttclid",
        "referrer",
        "landing",
    ):
        val = raw.get(key)
        if val is None:
            continue
        s = str(val).strip()
        if not s:
            continue
        out[key] = s[:500]
    return out


def _is_paid_medium(medium: str) -> bool:
    m = (medium or "").lower().strip()
    if not m:
        return False
    if m.startswith("paid"):
        return True
    return m in (
        "cpc",
        "ppc",
        "paid",
        "paidsocial",
        "paid_social",
        "paid-social",
        "display",
        "ads",
        "ad",
        "cpm",
        "cpa",
    )


def classify_traffic_source(attr: Optional[dict[str, Any]]) -> str:
    """
    Rotulo curto para o dashboard:
      Google Ads | Meta | Organico | Direto | Outro
    """
    a = normalize_attribution(attr)
    utm_source = (a.get("utm_source") or "").lower()
    utm_medium = (a.get("utm_medium") or "").lower()
    gclid = a.get("gclid") or ""
    fbclid = a.get("fbclid") or ""
    msclkid = a.get("msclkid") or ""
    ttclid = a.get("ttclid") or ""
    referrer = (a.get("referrer") or "").lower()
    paid = _is_paid_medium(utm_medium)

    # 1) Clique pago Google
    if gclid:
        return "Google Ads"
    if utm_source in ("googleads", "adwords", "google ads") or (
        "google" in utm_source and paid
    ):
        return "Google Ads"

    # 2) Meta (Facebook / Instagram)
    meta_keys = ("facebook", "fb", "instagram", "ig", "meta")
    if fbclid or utm_source in ("facebook", "fb", "instagram", "ig", "meta", "an", "fbads"):
        return "Meta"
    if any(k in utm_source for k in meta_keys):
        return "Meta"

    # 3) Outros pagos
    if msclkid or ttclid or (paid and utm_source):
        return "Outro"

    # 4) UTM organico / newsletter
    if utm_source:
        if "google" in utm_source:
            return "Organico"
        return "Outro"

    # 5) Referrer
    host = ""
    try:
        host = (urlparse(referrer).hostname or "").lower()
    except Exception:
        host = ""

    if host and any(
        h in host
        for h in (
            "google.",
            "bing.",
            "duckduckgo.",
            "yahoo.",
            "ecosia.",
            "search.brave.",
            "busca.uol.",
        )
    ):
        return "Organico"

    if host and any(
        h in host
        for h in (
            "facebook.",
            "fb.com",
            "instagram.",
            "l.facebook.",
            "lm.facebook.",
            "m.facebook.",
        )
    ):
        return "Meta"

    if not referrer:
        return "Direto"

    return "Outro"
