# -*- coding: utf-8 -*-
"""Locale / brand detection for Radar (PT) vs Winner Tracker (EN)."""

from __future__ import annotations

from typing import Any, Optional

EN_HOST_SUFFIXES = (
    "winnertracker.com",
)

BRAND_PT = {
    "locale": "pt",
    "brand": "Radar da Concorrência",
    "brand_short": "Radar",
    "mark": "R",
    "tagline": "Consultor de Marketing",
}

BRAND_EN = {
    "locale": "en",
    "brand": "Winner Tracker",
    "brand_short": "Winner Tracker",
    "mark": "W",
    "tagline": "Competitive Marketing Analyst",
}

STEP_LABELS_EN = {
    "briefing_concorrentes": "Briefing + competitors",
    "google_ads": "Google Ads",
    "seo": "Organic SEO",
    "brand": "Brand Search",
    "meta": "Meta Ads",
    "instagram": "Instagram",
    "youtube": "YouTube",
    "tiktok": "TikTok",
}


def normalize_locale(raw: Optional[str]) -> str:
    s = (raw or "").strip().lower()
    if s.startswith("en"):
        return "en"
    return "pt"


def host_implies_en(host: str) -> bool:
    h = (host or "").split(":")[0].strip().lower()
    if h.startswith("www."):
        h = h[4:]
    for suffix in EN_HOST_SUFFIXES:
        if h == suffix or h.endswith("." + suffix):
            return True
    return False


def resolve_locale(
    host: str = "",
    *,
    query_lang: str = "",
    body_locale: str = "",
) -> str:
    """Priority: explicit body/query > hostname > pt default."""
    for raw in (body_locale, query_lang):
        s = (raw or "").strip().lower()
        if s.startswith("en"):
            return "en"
        if s.startswith("pt"):
            return "pt"
    if host_implies_en(host):
        return "en"
    return "pt"


def brand_for(locale: str) -> dict[str, str]:
    return dict(BRAND_EN if normalize_locale(locale) == "en" else BRAND_PT)


def localize_steps(steps: list[dict[str, Any]], locale: str) -> list[dict[str, Any]]:
    if normalize_locale(locale) != "en":
        return steps
    out = []
    for s in steps:
        item = dict(s)
        lid = str(item.get("id") or "")
        if lid in STEP_LABELS_EN:
            item["label"] = STEP_LABELS_EN[lid]
        out.append(item)
    return out


def hello_payload(locale: str, *, live_ok: bool, missing_keys: Optional[list[str]] = None) -> dict[str, Any]:
    brand = brand_for(locale)
    missing = missing_keys or []
    if normalize_locale(locale) == "en":
        greeting = (
            f"Hi! I'm the AI agent **{brand['brand']}**. "
            "I analyze your company's marketing channels and your competitors, "
            "and show where you're winning or losing share and budget in your niche."
        )
        ask = (
            "To start the free analysis, send your **website URL** and your "
            "**phone number with country code** (same message is fine)."
        )
        examples = ["yourwebsite.com, +15551234567"]
        if not live_ok and missing:
            ask = (
                "To start the free analysis, send your **website URL** and your "
                f"**phone with country code**. (Server still missing API keys: {', '.join(missing)}.)"
            )
        return {
            **brand,
            "greeting": greeting,
            "ask": ask,
            "examples": examples,
        }

    greeting = (
        f"Olá! Eu sou o Agente de IA **{brand['brand']}**. "
        "Eu posso analisar os canais de marketing da sua empresa e dos "
        "seus concorrentes, e te mostrar onde você está ganhando ou perdendo "
        "espaço e dinheiro no seu nicho."
    )
    ask = (
        "Para iniciarmos a análise gratuita, envie o **endereço do seu site** e o seu "
        "**WhatsApp com DDD** (pode ser na mesma mensagem)."
    )
    examples = ["seusite.com.br, 11999999999"]
    if not live_ok and missing:
        ask = (
            "Para iniciarmos a análise gratuita, envie o **endereço do seu site** e o seu "
            f"**WhatsApp com DDD**. (Servidor ainda sem todas as API keys: {', '.join(missing)}.)"
        )
    return {
        **brand,
        "greeting": greeting,
        "ask": ask,
        "examples": examples,
    }


def chat_errors(locale: str) -> dict[str, str]:
    if normalize_locale(locale) == "en":
        return {
            "empty": "Empty message",
            "need_site": (
                "I couldn't identify the website. Send the URL and your phone with country code — "
                "e.g. `yourwebsite.com, +15551234567`."
            ),
            "need_phone": (
                "To unlock the free analysis, I need your **phone number with country code**. "
                "You can send it with the site, e.g. `yourwebsite.com, +15551234567`."
            ),
            "demo_only": (
                "Live analysis is not enabled in this environment yet. "
                "Set DEMO_ONLY=0 + API keys to analyze your site."
            ),
            "missing_keys": "Live analysis needs these API keys in Render Environment: ",
        }
    return {
        "empty": "Mensagem vazia",
        "need_site": (
            "Não identifiquei o site. Envie o endereço e o WhatsApp com DDD — "
            "ex: `seusite.com.br, 11999999999`."
        ),
        "need_phone": (
            "Para liberar a análise gratuita, preciso do seu **WhatsApp com DDD**. "
            "Pode enviar junto com o site, ex: `seusite.com.br, 11999999999`."
        ),
        "demo_only": (
            "Neste ambiente a analise ao vivo ainda nao esta liberada. "
            "Configure DEMO_ONLY=0 + API keys para analisar o seu site."
        ),
        "missing_keys": "Analise ao vivo precisa das API keys no Render Environment: ",
    }
