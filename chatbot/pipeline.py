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

try:
    import usage_db
except ImportError:
    usage_db = None  # type: ignore

CHATBOT_DIR = Path(__file__).resolve().parent
RADAR_ROOT = CHATBOT_DIR.parent
FINAL_DIR = RADAR_ROOT / "scripts"
SCRIPT_ENTENDER = FINAL_DIR / "1- entender o cliente.py"
SCRIPT_CONCORRENTES = FINAL_DIR / "3 - concorrentes Geral.py"
SCRIPT_GOOGLE_ADS = FINAL_DIR / "5a - google ads.py"
SCRIPT_SEO = FINAL_DIR / "5b - seo organico.py"
SCRIPT_BRAND = FINAL_DIR / "5c - brand search.py"
SCRIPT_META = FINAL_DIR / "5d - meta ads.py"
SCRIPT_LINKEDIN = FINAL_DIR / "5e - linkedin ads.py"
SCRIPT_SOCIAL = FINAL_DIR / "5f - social ig yt.py"
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

# Canais extra (apos free): Meta, LinkedIn, Instagram, YouTube — sem TikTok
EXTRA_STEP_DEFS = [
    {
        "id": "meta",
        "label": "Meta Ads",
        "tag": "1/4",
        "tag_cls": "amber",
        "index": 1,
        "eta_live": 180,
        "eta_demo": 12,
        "eta_cache": 10,
    },
    {
        "id": "linkedin",
        "label": "LinkedIn Ads",
        "tag": "2/4",
        "tag_cls": "blue",
        "index": 2,
        "eta_live": 180,
        "eta_demo": 12,
        "eta_cache": 10,
    },
    {
        "id": "instagram",
        "label": "Instagram",
        "tag": "3/4",
        "tag_cls": "green",
        "index": 3,
        "eta_live": 120,
        "eta_demo": 12,
        "eta_cache": 10,
    },
    {
        "id": "youtube",
        "label": "YouTube",
        "tag": "4/4",
        "tag_cls": "green",
        "index": 4,
        "eta_live": 30,
        "eta_demo": 10,
        "eta_cache": 9,
    },
]

EXTRA_TOTAL_STEPS = len(EXTRA_STEP_DEFS)

REVEAL_PAUSE = {
    "briefing_concorrentes": 1.6,
    "google_ads": 1.6,
    "seo": 1.5,
    "brand": 1.4,
    "meta": 1.2,
    "linkedin": 1.2,
    "instagram": 1.1,
    "youtube": 1.0,
}

# Demo publica: revelacao mais rapida
if __import__("os").environ.get("DEMO_ONLY", "").strip().lower() in ("1", "true", "yes"):
    REVEAL_PAUSE = {
        "briefing_concorrentes": 0.6,
        "google_ads": 0.7,
        "seo": 0.6,
        "brand": 0.6,
        "meta": 0.5,
        "linkedin": 0.5,
        "instagram": 0.5,
        "youtube": 0.5,
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


def _step_meta(step_id: str, defs: Optional[list] = None) -> Optional[dict]:
    pool = defs if defs is not None else (STEP_DEFS + EXTRA_STEP_DEFS)
    return next((s for s in pool if s["id"] == step_id), None)


def _step_eta_sec(step_id: str, mode: str, defs: Optional[list] = None) -> int:
    meta = _step_meta(step_id, defs)
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
    *,
    defs: Optional[list] = None,
    total: Optional[int] = None,
) -> None:
    active_defs = defs if defs is not None else STEP_DEFS
    meta = _step_meta(step_id, active_defs)
    index = int(meta["index"]) if meta else 0
    total_n = total if total is not None else len(active_defs)
    eta_sec = _step_eta_sec(step_id, mode, active_defs)
    set_step(step_id, state, detail)
    emit(
        "progress",
        step_id=step_id,
        state=state,
        detail=detail,
        index=index,
        total=total_n,
        label=meta["label"] if meta else step_id,
        eta_seconds=eta_sec,
        eta_label=_fmt_eta(eta_sec),
        progress_label=f"Etapa {index}/{total_n}",
    )


def _record_competitors_step(run_id: Optional[str], early: dict, step_id: str = "briefing_concorrentes") -> None:
    if not run_id or not usage_db:
        return
    try:
        meta = early.get("filter_meta") or {}
        stats = early.get("competitors_stats") or meta.get("counts") or {}
        usage_db.upsert_step(
            run_id,
            step_id,
            "done",
            f"{early.get('competitors_count', 0)} concorrentes na UI "
            f"(tier={early.get('display_tier') or meta.get('display_tier')})",
            companies_found=int(stats.get("total") or meta.get("companies_found") or 0),
            altos_found=int(stats.get("alto") or meta.get("altos_found") or 0),
            medios_found=int(stats.get("medio") or meta.get("medios_found") or 0),
            baixos_found=int(stats.get("baixo") or meta.get("baixos_found") or 0),
            competitors_count_ui=int(early.get("competitors_count") or 0),
            display_tier=str(early.get("display_tier") or meta.get("display_tier") or ""),
            meta={
                "note": early.get("competitors_note") or "",
                "stats": stats,
            },
        )
        raw = early.get("competitors_raw") or []
        if raw:
            # serializable snapshot
            serial = []
            for c in raw:
                if isinstance(c, dict):
                    serial.append({
                        k: c.get(k)
                        for k in (
                            "domain", "name", "similaridade", "perfil", "fonte",
                            "nicho", "url", "is_client",
                        )
                    })
            usage_db.save_json_artifact(
                run_id,
                "competitors_raw.json",
                serial,
                kind="competitors_raw",
                step_id=step_id,
                meta={"count": len(serial)},
            )
        usage_db.save_json_artifact(
            run_id,
            "competitors_ui.json",
            early.get("competitors") or [],
            kind="competitors_ui",
            step_id=step_id,
        )
        if early.get("competitors_note"):
            usage_db.append_log(
                run_id,
                early["competitors_note"],
                level="warn",
                source="competitors",
            )
    except Exception as e:
        try:
            usage_db.append_log(run_id, f"Falha ao gravar concorrentes: {e}", level="error")
        except Exception:
            pass


def _record_xlsx_artifacts(run_id: Optional[str], slug: str, step_id: str) -> None:
    if not run_id or not usage_db:
        return
    try:
        for kind, finder in (
            ("briefing_xlsx", find_briefing_xlsx),
            ("concorrentes_xlsx", find_concorrentes_xlsx),
            ("metricas_xlsx", find_metricas_xlsx),
        ):
            path = finder(slug)
            if path and path.exists():
                usage_db.add_artifact(
                    run_id,
                    kind=kind,
                    path=str(path),
                    step_id=step_id,
                    meta={"name": path.name},
                )
    except Exception:
        pass


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
            "competitors_note": report.get("competitors_note") or "",
            "display_tier": report.get("display_tier") or "",
            "competitors_stats": report.get("competitors_stats") or {},
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
    elif section in ("meta", "linkedin", "instagram", "youtube"):
        payload["data"] = report.get(section) or {}
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
    run_id: Optional[str] = None,
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
    _record_competitors_step(run_id, report)
    _record_xlsx_artifacts(run_id, slug, "briefing_concorrentes")
    return _reveal_from_report(report, emit, set_step, mode)


def run_live_pipeline(
    parsed: ParsedInput,
    emit: EmitFn,
    set_step: EmitFn,
    run_id: Optional[str] = None,
) -> dict:
    import os

    required = (
        "OPENROUTER_API_KEY",
        "SCRAPINGBEE_API_KEY",
        "DATAFORSEO_USER",
        "DATAFORSEO_PASS",
        "SEMRUSH_API_KEY",
    )
    missing = [k for k in required if not os.environ.get(k, "").strip()]
    if missing:
        raise RuntimeError(
            "API keys ausentes para live: " + ", ".join(missing)
        )

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
        _record_competitors_step(run_id, report)
        _record_xlsx_artifacts(run_id, slug, "briefing_concorrentes")
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
    _record_competitors_step(run_id, early)
    _record_xlsx_artifacts(run_id, slug, "briefing_concorrentes")
    _emit_progress(
        emit, set_step, "briefing_concorrentes", "done",
        f"{early.get('competitors_count', 0)} concorrentes mapeados"
        + (f" ({early.get('display_tier')})" if early.get("display_tier") else ""),
        mode,
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
    _record_xlsx_artifacts(run_id, slug, "brand")

    return report


def run_pipeline(
    parsed: ParsedInput,
    emit: EmitFn,
    set_step: EmitFn,
    run_id: Optional[str] = None,
) -> dict:
    if parsed.demo:
        return run_demo_pipeline(parsed, emit, set_step, run_id=run_id)
    return run_live_pipeline(parsed, emit, set_step, run_id=run_id)


def run_extras_pipeline(
    *,
    slug: str,
    company: str,
    emit: EmitFn,
    set_step: EmitFn,
    demo: bool = False,
    preferred_competitors: Optional[list[str]] = None,
    run_id: Optional[str] = None,
) -> dict:
    """Canais extra: Meta → LinkedIn → Instagram → YouTube (sem TikTok)."""
    slug = slug or "chatguru"
    client_name = company or slug
    mode = "demo" if demo else "live"
    defs = EXTRA_STEP_DEFS
    total = EXTRA_TOTAL_STEPS

    def on_log(line: str) -> None:
        emit("log", line=line)

    first_eta = _step_eta_sec("meta", mode, defs)
    emit(
        "pipeline_meta",
        mode=mode,
        eta_seconds=first_eta,
        eta_label=_fmt_eta(first_eta),
        total_steps=total,
        extras=True,
    )

    existing = find_metricas_xlsx(slug)
    if demo:
        xlsx = existing if (existing and existing.exists()) else DEMO_XLSX
        if not xlsx.exists():
            raise FileNotFoundError(f"XLSX demo nao encontrado para extras: {xlsx}")
        report = build_report_from_xlsx(
            xlsx,
            client_name=client_name,
            preferred_competitors=preferred_competitors,
        )
        emit("log", line=f"[EXTRAS/DEMO] Revelando canais extra de {xlsx.name}")
        for sid, label in (
            ("meta", "Meta Ads"),
            ("linkedin", "LinkedIn Ads"),
            ("instagram", "Instagram"),
            ("youtube", "YouTube"),
        ):
            _emit_progress(
                emit, set_step, sid, "running", f"Ranking {label}...", mode,
                defs=defs, total=total,
            )
            time.sleep(REVEAL_PAUSE.get(sid, 0.6))
            _emit_section(emit, sid, report, client_name)
            _emit_progress(
                emit, set_step, sid, "done", f"{label} pronto", mode,
                defs=defs, total=total,
            )
        return report

    if not SCRIPT_META.exists() or not SCRIPT_LINKEDIN.exists() or not SCRIPT_SOCIAL.exists():
        raise FileNotFoundError("Scripts de canais extra (5d/5e/5f) nao encontrados.")

    # Meta
    _emit_progress(
        emit, set_step, "meta", "running", "Coletando Meta Ads Library...", mode,
        defs=defs, total=total,
    )
    _run_script([sys.executable, str(SCRIPT_META), slug], FINAL_DIR, on_log)
    xlsx = find_metricas_xlsx(slug)
    if not xlsx:
        raise FileNotFoundError(f"XLSX de metricas nao encontrado para '{slug}' (rode a analise gratuita antes).")
    report = build_report_from_xlsx(xlsx, client_name=client_name, preferred_competitors=preferred_competitors)
    _emit_section(emit, "meta", report, client_name)
    _emit_progress(emit, set_step, "meta", "done", "Meta Ads pronto", mode, defs=defs, total=total)

    # LinkedIn
    _emit_progress(
        emit, set_step, "linkedin", "running", "Coletando LinkedIn Ads...", mode,
        defs=defs, total=total,
    )
    _run_script([sys.executable, str(SCRIPT_LINKEDIN), slug], FINAL_DIR, on_log)
    xlsx = find_metricas_xlsx(slug) or xlsx
    report = build_report_from_xlsx(xlsx, client_name=client_name, preferred_competitors=preferred_competitors)
    _emit_section(emit, "linkedin", report, client_name)
    _emit_progress(emit, set_step, "linkedin", "done", "LinkedIn Ads pronto", mode, defs=defs, total=total)

    # Social = IG + YT (um script, duas revelacoes)
    _emit_progress(
        emit, set_step, "instagram", "running", "Coletando Instagram + YouTube...", mode,
        defs=defs, total=total,
    )
    _run_script([sys.executable, str(SCRIPT_SOCIAL), slug], FINAL_DIR, on_log)
    xlsx = find_metricas_xlsx(slug) or xlsx
    report = build_report_from_xlsx(xlsx, client_name=client_name, preferred_competitors=preferred_competitors)
    _emit_section(emit, "instagram", report, client_name)
    _emit_progress(emit, set_step, "instagram", "done", "Instagram pronto", mode, defs=defs, total=total)

    _emit_progress(
        emit, set_step, "youtube", "running", "Montando ranking YouTube...", mode,
        defs=defs, total=total,
    )
    time.sleep(REVEAL_PAUSE.get("youtube", 1.0))
    _emit_section(emit, "youtube", report, client_name)
    _emit_progress(emit, set_step, "youtube", "done", "YouTube pronto", mode, defs=defs, total=total)

    return report
