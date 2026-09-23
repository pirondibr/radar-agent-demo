# -*- coding: utf-8 -*-
"""Extrai empresa (URL/nome) e concorrentes opcionais da mensagem do cliente."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from urllib.parse import urlparse


@dataclass
class ParsedInput:
    company: str = ""
    url: str = ""
    slug: str = ""
    competitors: list[str] = field(default_factory=list)
    demo: bool = False
    raw: str = ""
    kind: str = ""  # "" | "extras"


def slugify_client(client_input: str) -> str:
    raw = (client_input or "").strip().lower()
    raw = raw.replace("https://", "").replace("http://", "").replace("www.", "").split("/")[0]
    return re.sub(r"[^a-z0-9]+", "", raw.split(".")[0]) or "cliente"


def normalize_url(raw: str) -> str:
    text = (raw or "").strip()
    if not text:
        return ""
    if not text.startswith(("http://", "https://")):
        if "." in text and " " not in text:
            text = "https://" + text
        else:
            return ""
    return text.rstrip("/") + "/"


def _extract_urls(text: str) -> list[str]:
    return re.findall(r"https?://[^\s,;]+|(?:www\.)?[a-z0-9][-a-z0-9.]*\.[a-z]{2,}(?:/[^\s,;]*)?", text, re.I)


def _split_competitors(chunk: str) -> list[str]:
    parts = re.split(r"[,;/]| e | ou |\n", chunk, flags=re.I)
    out: list[str] = []
    for p in parts:
        name = re.sub(r"\s+", " ", p).strip(" .:-")
        if not name or len(name) < 2:
            continue
        if name.lower() in {"nenhum", "nenhuma", "nao", "não", "n/a", "na"}:
            continue
        out.append(name)
    return out[:3]


def parse_user_message(message: str) -> ParsedInput:
    raw = (message or "").strip()
    result = ParsedInput(raw=raw)

    low = raw.lower().strip()
    if low in {"demo", "demo chatguru", "chatguru demo"} or low.startswith("demo "):
        result.demo = True
        result.company = "Chatguru"
        result.url = "https://chatguru.com.br/"
        result.slug = "chatguru"
        return result

    # Explicit competitor section: "concorrentes: a, b" or "vs a, b"
    comps: list[str] = []
    company_part = raw
    m = re.search(
        r"(?:concorrentes?|competitors?|vs\.?|versus)\s*[:\-]?\s*(.+)$",
        raw,
        re.I | re.S,
    )
    if m:
        comps = _split_competitors(m.group(1))
        company_part = raw[: m.start()].strip(" ,.;\n")

    urls = _extract_urls(company_part) or _extract_urls(raw)
    if urls:
        first = urls[0]
        result.url = normalize_url(first)
        result.slug = slugify_client(result.url or first)
        host = urlparse(result.url).netloc if result.url else first
        host = host.replace("www.", "").split("/")[0]
        result.company = host.split(".")[0].capitalize() if host else result.slug

    if not result.company:
        # First line / first token as company name
        line = company_part.split("\n")[0].strip()
        line = re.sub(r"^(empresa|cliente|site|url)\s*[:\-]?\s*", "", line, flags=re.I)
        # If still has competitor keywords leftover
        line = re.split(r"(?:concorrentes?|vs\.?)", line, flags=re.I)[0].strip(" ,.;")
        if line:
            result.company = line[:80]
            result.slug = slugify_client(line)
            if "." in line and " " not in line:
                result.url = normalize_url(line)

    if not comps:
        # Fallback: comma-separated after company on same line
        # e.g. "chatguru, blip, huggy, chatpro"
        tokens = [t.strip() for t in re.split(r"[,;]", raw) if t.strip()]
        if len(tokens) >= 2 and not m:
            if not result.company:
                result.company = tokens[0]
                result.slug = slugify_client(tokens[0])
                if "." in tokens[0] and " " not in tokens[0]:
                    result.url = normalize_url(tokens[0])
            comps = tokens[1:4]

    result.competitors = comps

    # Reuse existing chatguru output as demo when user asks for chatguru without forcing live
    if result.slug == "chatguru" and low in {"chatguru", "https://chatguru.com.br", "https://chatguru.com.br/", "chatguru.com.br"}:
        result.demo = True
        result.url = result.url or "https://chatguru.com.br/"
        result.company = result.company or "Chatguru"

    if not result.url and result.slug and result.slug != "cliente":
        # Placeholder URL so script 1 can try scrape by domain guess
        result.url = f"https://{result.slug}.com.br/"

    return result
