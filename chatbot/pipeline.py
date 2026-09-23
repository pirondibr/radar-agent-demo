# -*- coding: utf-8 -*-
"""Orquestra scripts Final 09 (1→3→5) com revelacao progressiva por etapa."""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path
from typing import Callable, Optional

from parse_input import ParsedInput, normalize_url, slugify_client
from report_builder import (
    build_early_briefing_competitors,
    build_report_from_xlsx,
)

FINAL_DIR = Path(r"C:\Users\Usuario\Desktop\Radar 09 2026\scripts")
SCRIPT_ENTENDER = FINAL_DIR / "1- entender o cliente.py"
SCRIPT_CONCORRENTES = FINAL_DIR / "3 - concorrentes Geral.py"
SCRIPT_GOOGLE_ADS = FINAL_DIR / "5a - google ads.py"
SCRIPT_SEO = FINAL_DIR / "5b - seo organico.py"
SCRIPT_BRAND = FINAL_DIR / "5c - brand search.py"
RADAR_ROOT = FINAL_DIR.parent
METRICAS_DIR = RADAR_ROOT / "outputs" / "metricas"
BRIEFING_DIR = RADAR_ROOT / "outputs" / "entender"
CONCORRENTES_DIR = RADAR_ROOT / "outputs" / "concorrentes"
DEMO_XLSX = METRICAS_DIR / "chatguru" / "metricas-concorrentes-chatguru.xlsx"

EmitFn = Callable[..., None]

# 4 etapas visuais (briefing + concorrentes unificados)
STEP_DEFS = [
    {
        "id": "briefing_concorrentes",
        "label": "Briefing + concorrentes",
        "tag": "1/4",
        "tag_cls": "green",
        "index": 1,
        "eta_live": 120,
        "eta_demo": 10,
        "eta_cache": 8,
    },
    {
        "id": "google_ads",
        "label": "Google Ads",
        "tag": "2/4",
        "tag_cls": "amber",
        "index": 2,
        "eta_live": 240,
        "eta_demo": 10,
        "eta_cache": 9,
    },
    {
        "id": "seo",
        "label": "SEO organico",
        "tag": "3/4",
        "tag_cls": "blue",
        "index": 3,
        "eta_live": 20,
        "eta_demo": 10,
        "eta_cache": 9,
    },
    {
        "id": "brand",
        "label": "Marca (Brand Search)",
        "tag": "4/4",
        "tag_cls": "green",
        "index": 4,
        "eta_live": 20,
        "eta_demo": 10,
        "eta_cache": 9,
    },
]

TOTAL_STEPS = len(STEP_DEFS)

REVEAL_PAUSE = {
    "briefing_concorrentes": 1.6,
    "google_ads": 1.6,
    "seo": 1.5,
    "brand": 1.4,
}


def find_metricas_xlsx(slug: str) -> Optional[Path]:
    folder = METRICAS_DIR / slug
    if not folder.exists():
        return None
    files = sorted(
        folder.glob(f"metricas-concorrentes-{slug}*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not files:
        files = sorted(folder.glob("metricas-concorrentes-*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def find_briefing_xlsx(slug: str) -> Optional[Path]:
    folder = BRIEFING_DIR / slug
    if not folder.exists():
        return None
    files = sorted(
        folder.glob(f"briefing-cliente-{slug}*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not files:
        files = sorted(folder.glob("briefing-cliente-*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def find_concorrentes_xlsx(slug: str) -> Optional[Path]:
    folder = CONCORRENTES_DIR / slug
    if not folder.exists():
        return None
    # Prefer nacional (pipeline live)
    nacional = sorted(
        folder.glob(f"concorrentes-all-{slug}-nacional*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if nacional:
        return nacional[0]
    files = sorted(
        folder.glob(f"concorrentes-all-{slug}*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not files:
        files = sorted(folder.glob("concorrentes-all-*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def _run_script(cmd: list[str], cwd: Path, on_log: Optional[Callable[[str], None]] = None) -> None:
    proc = subprocess.Popen(
        cmd,
        cwd=str(cwd),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    assert proc.stdout is not None
    for line in proc.stdout:
        line = line.rstrip("\n\r")
        if line and on_log:
            on_log(line)
    proc.wait()
    if proc.returncode != 0:
        raise RuntimeError(f"Comando falhou (codigo {proc.returncode}): {' '.join(cmd)}")


def _fmt_eta(seconds: int) -> str:
    if seconds >= 60:
        m = max(1, round(seconds / 60))
        return f"~{m} min"
    return f"~{seconds}s"


def _step_eta_sec(step_id: str, mode: str) -> int:
    meta = next((s for s in STEP_DEFS if s["id"] == step_id), None)
    if not meta:
        return 60
    key = f"eta_{mode}" if mode in ("live", "demo", "cache") else "eta_live"
    return int(meta.get(key) or meta.get("eta_live") or 60)


def _emit_progress(
    emit: EmitFn,
    set_step: EmitFn,
    step_id: str,
    state: str,
    detail: str,
    mode: str,
) -> None:
    meta = next((s for s in STEP_DEFS if s["id"] == step_id), None)
    index = int(meta["index"]) if meta else 0
    eta_sec = _step_eta_sec(step_id, mode)
    set_step(step_id, state, detail)
    emit(
        "progress",
        step_id=step_id,
        state=state,
        detail=detail,
        index=index,
        total=TOTAL_STEPS,
        label=meta["label"] if meta else step_id,
        eta_seconds=eta_sec,
        eta_label=_fmt_eta(eta_sec),
        progress_label=f"Etapa {index}/{TOTAL_STEPS}",
    )


def _emit_section(emit: EmitFn, section: str, report: dict, client: str = "") -> None:
    """Envia pedaco do relatorio para a UI renderizar na hora."""
    payload: dict = {"section": section, "client": report.get("client") or client}
    if section == "briefing_concorrentes":
        payload["data"] = {
            "client": report.get("client") or client,
            "briefing": report.get("briefing") or {
                "client": client,
                "url": "",
                "nicho": "",
                "escopo": "",
                "summary": f"Briefing de {client} pronto.",
            },
            "competitors": report.get("competitors") or [],
            "competitors_count": report.get("competitors_count") or 0,
        }
    elif section == "briefing":
        payload["data"] = report.get("briefing") or {
            "client": client,
            "url": "",
            "nicho": "",
            "escopo": "",
            "summary": f"Briefing de {client} pronto.",
        }
    elif section == "competitors":
        payload["data"] = {
            "client": report.get("client"),
            "competitors": report.get("competitors") or [],
            "competitors_count": report.get("competitors_count") or 0,
        }
    elif section == "google_ads":
        payload["data"] = report.get("google_ads") or {}
        payload["client"] = report.get("client")
    elif section == "seo":
        payload["data"] = report.get("seo") or {}
        payload["client"] = report.get("client")
    elif section == "brand":
        payload["data"] = report.get("brand") or {}
        payload["client"] = report.get("client")
    emit("partial", **payload)


def _reveal_from_report(
    report: dict,
    emit: EmitFn,
    set_step: EmitFn,
    mode: str,
    pauses: Optional[dict[str, float]] = None,
) -> dict:
    """Revela as 4 etapas em sequencia a partir de um report ja montado."""
    pauses = pauses or REVEAL_PAUSE
    client = report.get("client") or "Cliente"

    # 1 Briefing + Concorrentes
    _emit_progress(
        emit, set_step, "briefing_concorrentes", "running",
        "Montando briefing e lista de concorrentes...", mode,
    )
    time.sleep(pauses.get("briefing_concorrentes", 1.0))
    _emit_section(emit, "briefing_concorrentes", report, client)
    _emit_progress(
        emit, set_step, "briefing_concorrentes", "done",
        "Briefing e concorrentes prontos", mode,
    )

    # 2 Google Ads
    _emit_progress(emit, set_step, "google_ads", "running", "Ranking Google Ads...", mode)
    time.sleep(pauses.get("google_ads", 1.0))
    _emit_section(emit, "google_ads", report, client)
    _emit_progress(emit, set_step, "google_ads", "done", "Google Ads pronto", mode)

    # 3 SEO
    _emit_progress(emit, set_step, "seo", "running", "Ranking SEO...", mode)
    time.sleep(pauses.get("seo", 1.0))
    _emit_section(emit, "seo", report, client)
    _emit_progress(emit, set_step, "seo", "done", "SEO pronto", mode)

    # 4 Marca
    _emit_progress(emit, set_step, "brand", "running", "Ranking de marca...", mode)
    time.sleep(pauses.get("brand", 1.0))
    _emit_section(emit, "brand", report, client)
    _emit_progress(emit, set_step, "brand", "done", "Marca pronta", mode)

    return report


def run_demo_pipeline(
    parsed: ParsedInput,
    emit: EmitFn,
    set_step: EmitFn,
) -> dict:
    slug = parsed.slug or "chatguru"
    xlsx = find_metricas_xlsx(slug) if slug != "chatguru" else DEMO_XLSX
    if xlsx is None or not xlsx.exists():
        xlsx = DEMO_XLSX
    if not xlsx.exists():
        raise FileNotFoundError(f"XLSX demo nao encontrado: {xlsx}")

    mode = "demo"
    first_eta = _step_eta_sec("briefing_concorrentes", mode)
    emit(
        "pipeline_meta",
        mode=mode,
        eta_seconds=first_eta,
        eta_label=_fmt_eta(first_eta),
        total_steps=TOTAL_STEPS,
    )
    emit("log", line=f"[DEMO] Usando dados prontos: {xlsx.name}")

    report = build_report_from_xlsx(
        xlsx,
        client_name=parsed.company or slug,
        preferred_competitors=parsed.competitors,
    )
    return _reveal_from_report(report, emit, set_step, mode)


def run_live_pipeline(
    parsed: ParsedInput,
    emit: EmitFn,
    set_step: EmitFn,
) -> dict:
    if not FINAL_DIR.exists():
        raise FileNotFoundError(f"Final 09 nao encontrado: {FINAL_DIR}")

    url = normalize_url(parsed.url) or parsed.url
    if not url:
        raise ValueError("URL do cliente obrigatoria para pipeline live.")
    slug = parsed.slug or slugify_client(url)
    client_name = parsed.company or slug

    def on_log(line: str) -> None:
        emit("log", line=line)

    existing = find_metricas_xlsx(slug)
    if existing and existing.exists():
        mode = "cache"
        first_eta = _step_eta_sec("briefing_concorrentes", mode)
        emit(
            "pipeline_meta",
            mode=mode,
            eta_seconds=first_eta,
            eta_label=_fmt_eta(first_eta),
            total_steps=TOTAL_STEPS,
        )
        emit("log", line=f"[CACHE] Metricas existentes para '{slug}': {existing.name}")
        report = build_report_from_xlsx(
            existing,
            client_name=client_name,
            preferred_competitors=parsed.competitors,
        )
        return _reveal_from_report(report, emit, set_step, mode)

    mode = "live"
    first_eta = _step_eta_sec("briefing_concorrentes", mode)
    emit(
        "pipeline_meta",
        mode=mode,
        eta_seconds=first_eta,
        eta_label=_fmt_eta(first_eta),
        total_steps=TOTAL_STEPS,
    )

    # --- Etapa 1/4: scripts 1 + 3, depois mostra briefing E concorrentes ---
    _emit_progress(
        emit, set_step, "briefing_concorrentes", "running",
        f"Briefing + mapeamento de concorrentes: {url}", mode,
    )
    if not SCRIPT_ENTENDER.exists():
        raise FileNotFoundError(SCRIPT_ENTENDER)
    _run_script([sys.executable, str(SCRIPT_ENTENDER), url], FINAL_DIR, on_log)

    if not SCRIPT_CONCORRENTES.exists():
        raise FileNotFoundError(SCRIPT_CONCORRENTES)
    _emit_progress(
        emit, set_step, "briefing_concorrentes", "running",
        "SEO + LLM: mapeando concorrentes nacionais...", mode,
    )
    _run_script(
        [sys.executable, str(SCRIPT_CONCORRENTES), slug, "nacional"],
        FINAL_DIR,
        on_log,
    )

    early = build_early_briefing_competitors(
        find_briefing_xlsx(slug),
        find_concorrentes_xlsx(slug),
        client_name=client_name,
        preferred_competitors=parsed.competitors,
        fallback_url=url,
    )
    _emit_section(emit, "briefing_concorrentes", early, client_name)
    _emit_progress(
        emit, set_step, "briefing_concorrentes", "done",
        f"{early.get('competitors_count', 0)} concorrentes mapeados", mode,
    )

    # --- Etapa 2/4: somente Google Ads ---
    _emit_progress(
        emit, set_step, "google_ads", "running",
        "Coletando Google Ads Transparency...", mode,
    )
    if not SCRIPT_GOOGLE_ADS.exists():
        raise FileNotFoundError(SCRIPT_GOOGLE_ADS)
    _run_script([sys.executable, str(SCRIPT_GOOGLE_ADS), slug], FINAL_DIR, on_log)

    xlsx = find_metricas_xlsx(slug)
    if not xlsx:
        raise FileNotFoundError(f"XLSX de metricas nao gerado para '{slug}'")

    report = build_report_from_xlsx(
        xlsx,
        client_name=client_name,
        preferred_competitors=parsed.competitors,
    )
    _emit_section(emit, "google_ads", report, client_name)
    _emit_progress(emit, set_step, "google_ads", "done", "Google Ads pronto", mode)

    # --- Etapa 3/4: somente SEO (Semrush do script 3) ---
    _emit_progress(emit, set_step, "seo", "running", "Montando ranking SEO organico...", mode)
    if not SCRIPT_SEO.exists():
        raise FileNotFoundError(SCRIPT_SEO)
    _run_script([sys.executable, str(SCRIPT_SEO), slug], FINAL_DIR, on_log)
    xlsx = find_metricas_xlsx(slug) or xlsx
    report = build_report_from_xlsx(
        xlsx,
        client_name=client_name,
        preferred_competitors=parsed.competitors,
    )
    _emit_section(emit, "seo", report, client_name)
    _emit_progress(emit, set_step, "seo", "done", "SEO pronto", mode)

    # --- Etapa 4/4: somente Brand Search ---
    _emit_progress(emit, set_step, "brand", "running", "Montando ranking de marca...", mode)
    if not SCRIPT_BRAND.exists():
        raise FileNotFoundError(SCRIPT_BRAND)
    _run_script([sys.executable, str(SCRIPT_BRAND), slug], FINAL_DIR, on_log)
    xlsx = find_metricas_xlsx(slug) or xlsx
    report = build_report_from_xlsx(
        xlsx,
        client_name=client_name,
        preferred_competitors=parsed.competitors,
    )
    _emit_section(emit, "brand", report, client_name)
    _emit_progress(emit, set_step, "brand", "done", "Marca pronta", mode)

    return report


def run_pipeline(parsed: ParsedInput, emit: EmitFn, set_step: EmitFn) -> dict:
    if parsed.demo:
        return run_demo_pipeline(parsed, emit, set_step)
    return run_live_pipeline(parsed, emit, set_step)
