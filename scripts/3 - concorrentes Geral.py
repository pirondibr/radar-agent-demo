# -*- coding: utf-8 -*-
"""Runner local de concorrentes por caminho/modo.

Salva em:
    Radar 09 2026/outputs/concorrentes/<cliente>/

Otimizacoes (modo nacional):
  - seo_depth=10 (alinhado ao consolidate)
  - SEO SERP e lista LLM em paralelo
  - uma unica passada de analyze_all_competitors (sem analisar 2x)

Uso:
    python "3 - concorrentes Geral.py" chatguru nacional
"""

from __future__ import annotations

import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

scripts_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(scripts_dir))

from workspace_paths import (  # noqa: E402
    CONC_DIR,
    OUT_CONCORRENTES,
    setup_workspace,
)

sys.path.insert(0, str(CONC_DIR))
sys.path.insert(0, str(scripts_dir))

setup_workspace()

from find_concorrentes import (  # noqa: E402
    MAX_KEYWORDS_KW,
    analyze_all_competitors,
    build_3period_traffic,
    consolidate_competitors,
    find_client_xlsx,
    keywords_from_briefing_seo,
    read_briefing_from_xlsx,
    read_keywords_from_xlsx,
    read_similarity_from_xlsx,
    select_keywords,
)
from find_concorrentes_all import (  # noqa: E402
    SERP_DEPTH_TOP10,
    find_competitors_via_llm,
    merge_competitors,
    run_ads_mode,
    run_llm_mode,
    run_maps_mode,
    run_seo_mode,
    write_unified_xlsx,
)
from gerar_html_concorrentes import generate_html_from_xlsx  # noqa: E402
from seo_pipeline import (  # noqa: E402
    brand_token_from_domain,
    classify_client_profile,
    fetch_serps,
    scrape_site,
)


LOCAL_OUT_DIR = OUT_CONCORRENTES

EXECUTION_MODES = {
    "nacional": {
        "label": "Nacional",
        "channels": ["SEO", "LLM"],
        "seo_depth": SERP_DEPTH_TOP10,  # 10 (antes 20; consolidate so usa top 10)
        "traffic_limit": 15,
        "fast": True,
    },
    "completo": {
        "label": "Completo",
        "channels": ["SEO", "Maps", "Ads", "LLM"],
        "seo_depth": SERP_DEPTH_TOP10,
        "traffic_limit": 15,
        "fast": False,
    },
}


def _select_seo_keywords(briefing: dict, xlsx_path: Path) -> list:
    selected = keywords_from_briefing_seo(briefing)
    if selected:
        for k in selected:
            k["source"] = "briefing SEO (entender o cliente)"
        print(f"[SEO/2] Usando {len(selected)} termo(s) SEO do briefing:")
        for k in selected:
            print(f"        - '{k['keyword']}' ({k['source']})")
        return selected

    print("[SEO/2] Sem termos SEO no briefing; fallback para Palavras-chave do XLSX")
    kws_list = read_keywords_from_xlsx(xlsx_path)
    sim_rows = read_similarity_from_xlsx(xlsx_path)
    client_brand = brand_token_from_domain(briefing.get("url", ""))
    selected = select_keywords(kws_list, sim_rows, brand=client_brand)
    for i, k in enumerate(selected):
        k["source"] = "palavras-chave (top vol)" if i < MAX_KEYWORDS_KW else "similaridade Alto"
    print(f"[SEO/2] {len(selected)} keyword(s) selecionada(s):")
    for k in selected:
        print(f"        - '{k['keyword']}' (vol {k.get('volume')}, {k['source']})")
    return selected


def _discover_seo_raw(briefing: dict, xlsx_path: Path, serp_depth: int) -> dict:
    print("\n" + "=" * 60)
    print(f" DISCOVERY SEO (SERP top {serp_depth}, sem analyze)")
    print("=" * 60)
    selected = _select_seo_keywords(briefing, xlsx_path)
    if not selected:
        return {"selected_kws": [], "competitors_raw": []}
    print(f"[SEO/3] Buscando SERPs top {serp_depth} ...")
    serps = fetch_serps([k["keyword"] for k in selected], depth=serp_depth, workers=5)
    competitors = consolidate_competitors(serps, briefing["url"])
    print(f"[SEO/3] {len(competitors)} concorrentes unicos no top {serp_depth}")
    return {"selected_kws": selected, "competitors_raw": competitors}


def _discover_llm_raw(briefing: dict) -> dict:
    print("\n" + "=" * 60)
    print(" DISCOVERY LLM (lista Gemini, sem analyze)")
    print("=" * 60)
    client_url = briefing.get("url", "")
    if not client_url:
        print("[LLM] URL do cliente ausente. Modo LLM cancelado.")
        return {"competitors_raw": []}
    competitors = find_competitors_via_llm(client_url, briefing)
    return {"competitors_raw": competitors or []}


def _dedupe_raw(seo_raw: list, llm_raw: list) -> tuple[list, set, set]:
    """Une listas cruas por dominio. Retorna (combined, seo_domains, llm_domains)."""
    by_dom: dict = {}
    seo_domains: set = set()
    llm_domains: set = set()

    for c in seo_raw:
        d = c.get("domain")
        if not d:
            continue
        seo_domains.add(d)
        by_dom[d] = dict(c)

    for c in llm_raw:
        d = c.get("domain")
        if not d:
            continue
        llm_domains.add(d)
        if d in by_dom:
            existing = by_dom[d]
            merged_kws = list(existing.get("keywords") or []) + list(c.get("keywords") or [])
            existing["keywords"] = merged_kws
            if not existing.get("title") and c.get("title"):
                existing["title"] = c["title"]
            if not existing.get("url") and c.get("url"):
                existing["url"] = c["url"]
        else:
            by_dom[d] = dict(c)

    return list(by_dom.values()), seo_domains, llm_domains


def run_nacional_fast(briefing: dict, xlsx_path: Path, mode_cfg: dict):
    """SEO ∥ LLM discovery, depois uma unica analise de similaridade."""
    serp_depth = int(mode_cfg.get("seo_depth") or SERP_DEPTH_TOP10)

    with ThreadPoolExecutor(max_workers=2) as pool:
        fut_seo = pool.submit(_discover_seo_raw, briefing, xlsx_path, serp_depth)
        fut_llm = pool.submit(_discover_llm_raw, briefing)
        seo_disc = fut_seo.result()
        llm_disc = fut_llm.result()

    combined_raw, seo_domains, llm_domains = _dedupe_raw(
        seo_disc.get("competitors_raw") or [],
        llm_disc.get("competitors_raw") or [],
    )
    print("\n" + "=" * 60)
    print(
        f" ANALYZE UNICO: {len(combined_raw)} dominios "
        f"(SEO={len(seo_domains)}, LLM={len(llm_domains)}, overlap="
        f"{len(seo_domains & llm_domains)})"
    )
    print("=" * 60)

    analyzed = []
    if combined_raw:
        analyzed = analyze_all_competitors(
            combined_raw, briefing["url"], briefing, workers=5,
        )

    seo_full = [c for c in analyzed if c.get("domain") in seo_domains]
    llm_full = [c for c in analyzed if c.get("domain") in llm_domains]

    seo_result = {
        "selected_kws": seo_disc.get("selected_kws") or [],
        "competitors_full": seo_full,
    }
    llm_result = {"competitors_full": llm_full}
    empty_maps = {"maps_lookup": {}, "competitors_full": []}
    empty_ads = {"keywords": [], "competitors_full": [], "raw_results": []}
    return seo_result, empty_maps, empty_ads, llm_result


def main() -> None:
    if len(sys.argv) < 2:
        print('Uso: python "3 - concorrentes Geral.py" <cliente> [modo]')
        print('Ex:  python "3 - concorrentes Geral.py" chatguru nacional')
        sys.exit(1)

    client_input = sys.argv[1].strip()
    mode_key = (sys.argv[2].strip().lower() if len(sys.argv) >= 3 else "nacional")
    if mode_key not in EXECUTION_MODES:
        print(f"ERRO: modo '{mode_key}' invalido. Opcoes: {', '.join(EXECUTION_MODES)}")
        sys.exit(1)

    mode_cfg = EXECUTION_MODES[mode_key]
    channels = set(mode_cfg["channels"])
    t0 = time.time()

    print(f"\n=== Concorrentes Radar: '{client_input}' | caminho {mode_cfg['label']} ===\n")

    xlsx_path = find_client_xlsx(client_input)
    briefing = read_briefing_from_xlsx(xlsx_path)
    if not briefing.get("url"):
        print("ERRO: URL do cliente nao encontrada no briefing.")
        sys.exit(1)
    client_slug = xlsx_path.parent.name

    if mode_cfg.get("fast") and channels == {"SEO", "LLM"}:
        seo_result, maps_result, ads_result, llm_result = run_nacional_fast(
            briefing, xlsx_path, mode_cfg,
        )
    else:
        seo_result = (
            run_seo_mode(briefing, xlsx_path, serp_depth=mode_cfg["seo_depth"])
            if "SEO" in channels else
            {"selected_kws": [], "competitors_full": []}
        )
        maps_result = (
            run_maps_mode(briefing, client_slug)
            if "Maps" in channels else
            {"maps_lookup": {}, "competitors_full": []}
        )
        ads_result = (
            run_ads_mode(briefing, seo_result.get("selected_kws") or [])
            if "Ads" in channels else
            {"keywords": [], "competitors_full": [], "raw_results": []}
        )
        llm_result = (
            run_llm_mode(briefing)
            if "LLM" in channels else
            {"competitors_full": []}
        )

    print("\n" + "=" * 60)
    print(f" MERGE: {' + '.join(mode_cfg['channels'])}")
    print("=" * 60)
    merged = merge_competitors(
        seo_result.get("competitors_full") or [],
        maps_result.get("competitors_full") or [],
        ads_result.get("competitors_full") or [],
        llm_result.get("competitors_full") or [],
    )
    print(f"[MERGE] {len(merged)} concorrentes unicos")

    client_url = briefing["url"]
    if not client_url.startswith(("http://", "https://")):
        client_url = "https://" + client_url
    try:
        site = scrape_site(client_url)
        client_profile = classify_client_profile(client_url, site, briefing)
    except Exception as e:
        print(f"[6] Falha ao classificar cliente: {e}")
        client_profile = "—"

    traffic = build_3period_traffic(
        client_url,
        merged,
        client_profile,
        max_competitors=mode_cfg["traffic_limit"],
    )

    suffix = "" if mode_key == "completo" else f"-{mode_key}"
    out_path = LOCAL_OUT_DIR / client_slug / f"concorrentes-all-{client_slug}{suffix}.xlsx"
    saved = write_unified_xlsx(
        out_path,
        client_input,
        briefing,
        seo_result,
        maps_result,
        ads_result,
        llm_result,
        merged,
        traffic,
        client_profile,
        mode_key=mode_key,
        mode_cfg=mode_cfg,
    )
    print(f"[OUT] XLSX salvo: {saved}")

    html_path = generate_html_from_xlsx(saved)
    print(f"[OUT] HTML salvo: {html_path}")
    print(f"\n=== Concluido em {time.time() - t0:.1f}s ===")


if __name__ == "__main__":
    main()
