# -*- coding: utf-8 -*-
"""Caminhos centralizados do Radar 09 2026.

Scripts em Radar/scripts/
Outputs em Radar/outputs/{entender,concorrentes,metricas}/
Libs Spy (SEO/LLM/find_concorrentes) permanecem referenciadas por caminho absoluto.
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
RADAR_ROOT = SCRIPTS_DIR.parent

SPY_ROOT = Path(r"C:\Users\Usuario\Desktop\Spy versao mes 09")
BASE_DIR = SPY_ROOT / "Lista automatica vertical 1"
CONC_DIR = BASE_DIR / "Concorrentes"
MVP_DIR = SPY_ROOT / "Mvp nova sequencia para Local"

PIPELINE_OUTPUT_DIR = RADAR_ROOT / "outputs" / "entender"
OUT_DIR = RADAR_ROOT / "outputs"
OUT_CONCORRENTES = RADAR_ROOT / "outputs" / "concorrentes"
OUT_METRICAS = RADAR_ROOT / "outputs" / "metricas"
FORMULA_DIR = Path(r"C:\Users\Usuario\Desktop\Radar Concorrencia\Formula de potencial de canais")

# Compat com codigo antigo que ainda espera FINAL_DIR = pasta dos runners
FINAL_DIR = SCRIPTS_DIR

SEO_PIPELINE_CANDIDATES = (
    BASE_DIR / "seo_pipeline.py",
    BASE_DIR / "seo_pipeline Antigo não usar.py",
)


def ensure_output_dirs() -> None:
    for d in (PIPELINE_OUTPUT_DIR, OUT_CONCORRENTES, OUT_METRICAS):
        d.mkdir(parents=True, exist_ok=True)


def setup_workspace() -> Path:
    """Prioriza scripts do Radar no sys.path e registra seo_pipeline."""
    ensure_output_dirs()

    # Radar scripts primeiro (ganha do Concorrentes/workspace_paths.py)
    if str(SCRIPTS_DIR) in sys.path:
        sys.path.remove(str(SCRIPTS_DIR))
    sys.path.insert(0, str(SCRIPTS_DIR))

    if str(CONC_DIR) not in sys.path:
        sys.path.insert(1, str(CONC_DIR))
    if str(BASE_DIR) not in sys.path:
        sys.path.insert(2, str(BASE_DIR))

    if "seo_pipeline" not in sys.modules:
        loaded = False
        for candidate in SEO_PIPELINE_CANDIDATES:
            if not candidate.exists():
                continue
            spec = importlib.util.spec_from_file_location("seo_pipeline", candidate)
            if spec is None or spec.loader is None:
                continue
            module = importlib.util.module_from_spec(spec)
            sys.modules["seo_pipeline"] = module
            spec.loader.exec_module(module)
            loaded = True
            break
        if not loaded:
            raise FileNotFoundError(
                "seo_pipeline nao encontrado. Esperado em:\n"
                + "\n".join(f"  - {p}" for p in SEO_PIPELINE_CANDIDATES)
            )

    return BASE_DIR
