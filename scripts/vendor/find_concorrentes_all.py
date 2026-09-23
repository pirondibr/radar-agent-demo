# -*- coding: utf-8 -*-
"""
Concorrentes Finder - modos de execucao (SEO / Maps / Ads / LLM)

Roda os quatro modos em sequencia e produz UM UNICO XLSX com:

  Aba 1: Resumo (info dos quatro modos)
  Aba 2: Concorrentes SEO          (top 10 SERP organico)
  Aba 3: Concorrentes Maps         (top 10 Google Maps na categoria)
  Aba 4: Concorrentes Ads          (anunciantes patrocinados do Google Ads)
  Aba 5: Concorrentes LLM          (concorrentes diretos sugeridos por Gemini)
  Aba 6: Concorrentes Unificado    (dedupe entre as 4 fontes + coluna Fonte)
  Aba 7: Trafego SEO Marca         (cliente + Alto unicos do unificado)

Uso:
    python find_concorrentes_all.py <nome_cliente> [modo]
    python find_concorrentes_all.py cnempreendimentos
    python find_concorrentes_all.py chatguru nacional
"""

import re
import sys
import time
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ----------------------------------------------------------------------
# Imports do pipeline
# ----------------------------------------------------------------------
_THIS_DIR = Path(__file__).resolve().parent
from workspace_paths import setup_workspace  # noqa: E402

setup_workspace()

from seo_pipeline import (  # type: ignore  # noqa: E402
    fetch_serps, normalize_domain, classify_client_profile, scrape_site,
    call_openrouter, extract_json,
)
from find_concorrentes import (  # type: ignore  # noqa: E402
    find_client_xlsx, read_briefing_from_xlsx, read_keywords_from_xlsx,
    read_similarity_from_xlsx, select_keywords, consolidate_competitors,
    analyze_all_competitors, build_3period_traffic,
    keywords_from_briefing_seo,
    SERP_DEPTH_TOP10, MAX_KEYWORDS_KW, FILL_SIM, FILL_PERFIL,
    HEAD_FONT, HEAD_FILL, BORDER,
)
from seo_pipeline import brand_token_from_domain  # type: ignore  # noqa: E402
from find_concorrentes_maps import (  # type: ignore  # noqa: E402
    find_client_in_maps, find_competitors_in_maps,
    consolidate_maps_competitors, derive_client_name, MAPS_TOP_N,
)

try:
    from scrape_ads_browser import research_ads_for_keywords  # type: ignore  # noqa: E402
    ADS_AVAILABLE = True
except ImportError:
    ADS_AVAILABLE = False

    def research_ads_for_keywords(keywords, **kwargs):  # type: ignore
        return []


OUT_DIR = _THIS_DIR / "output"

EXECUTION_MODES = {
    "completo": {
        "label": "Completo",
        "channels": ["SEO", "Maps", "Ads", "LLM"],
        "seo_depth": SERP_DEPTH_TOP10,
        "traffic_limit": 15,
    },
    "nacional": {
        "label": "Nacional",
        "channels": ["SEO", "LLM"],
        "seo_depth": 20,
        "traffic_limit": 15,
    },
}

# limite de keywords para o modo Ads (cada keyword toma ~20-45s com retry,
# entao 2 mantem o processo ~1min em vez de ~2min)
MAX_KEYWORDS_ADS = 2

# ignora dominios genericos que poluiriam o relatorio de Ads
ADS_DOMAINS_IGNORE = {
    "google.com", "googleadservices.com", "youtube.com", "facebook.com",
    "instagram.com", "linkedin.com",
}


# ----------------------------------------------------------------------
# Roda cada modo isoladamente, retornando dados estruturados
# ----------------------------------------------------------------------
def run_seo_mode(briefing: dict, xlsx_path: Path, serp_depth: int = SERP_DEPTH_TOP10) -> dict:
    """Retorna {selected_kws, competitors_full} do modo SEO orgânico.

    Prioridade das keywords:
      1) termos SEO do briefing (entender o cliente: "Defina 2 palavras para SEO")
      2) fallback: aba Palavras-chave / Similaridade do XLSX (relatorio-seo)
    """
    print("\n" + "=" * 60)
    print(f" MODO 1: SEO ORGANICO (Google SERP top {serp_depth})")
    print("=" * 60)

    selected = keywords_from_briefing_seo(briefing)
    if selected:
        for k in selected:
            k["source"] = "briefing SEO (entender o cliente)"
        print(f"[SEO/2] Usando {len(selected)} termo(s) SEO do briefing:")
        for k in selected:
            print(f"        - '{k['keyword']}' ({k['source']})")
    else:
        print("[SEO/2] Sem termos SEO no briefing; fallback para Palavras-chave do XLSX")
        kws_list = read_keywords_from_xlsx(xlsx_path)
        sim_rows = read_similarity_from_xlsx(xlsx_path)
        client_brand = brand_token_from_domain(briefing.get("url", ""))
        selected = select_keywords(kws_list, sim_rows, brand=client_brand)
        for i, k in enumerate(selected):
            k["source"] = "palavras-chave (top vol)" if i < MAX_KEYWORDS_KW else "similaridade Alto"

        print(f"[SEO/2] {len(selected)} keyword(s) selecionada(s):")
        for k in selected:
            print(f"        - '{k['keyword']}' (vol {k['volume']}, {k['source']})")

    if not selected:
        return {"selected_kws": [], "competitors_full": []}

    print(f"[SEO/3] Buscando SERPs top {serp_depth} ...")
    serps = fetch_serps([k["keyword"] for k in selected],
                         depth=serp_depth, workers=5)
    competitors = consolidate_competitors(serps, briefing["url"])
    print(f"[SEO/3] {len(competitors)} concorrentes unicos no top {serp_depth}")

    competitors_full = analyze_all_competitors(competitors, briefing["url"], briefing,
                                                 workers=5)
    return {"selected_kws": selected, "competitors_full": competitors_full}


def run_maps_mode(briefing: dict, client_slug: str) -> dict:
    """Retorna {maps_lookup, competitors_full} do modo Google Maps."""
    print("\n" + "=" * 60)
    print(" MODO 2: GOOGLE MAPS (categoria + raio)")
    print("=" * 60)

    client_name = derive_client_name(briefing, client_slug)
    print(f"[MAPS/2] Nome derivado: '{client_name}'")

    print("[MAPS/3a] Buscando cliente no Maps ...")
    maps_lookup = find_client_in_maps(client_name, briefing["url"], briefing)
    if not maps_lookup:
        print("[MAPS] Cliente nao encontrado no Maps. Modo Maps cancelado.")
        return {"maps_lookup": {}, "competitors_full": []}

    category = maps_lookup.get("category", "")
    lat = maps_lookup.get("latitude")
    lng = maps_lookup.get("longitude")
    print(f"[MAPS/3a] Cliente: '{maps_lookup.get('title','')}'")
    print(f"          categoria: {category}")
    print(f"          extras: {maps_lookup.get('additional_categories') or []}")
    print(f"          coord: {lat}, {lng}")
    print(f"          end: {maps_lookup.get('address','')}")

    if not category:
        print("[MAPS] Categoria nao identificada. Modo Maps cancelado.")
        return {"maps_lookup": maps_lookup, "competitors_full": []}

    cidades = briefing.get("cidades") or []
    cidade_principal = cidades[0] if cidades else ""
    print(f"\n[MAPS/3b] Top {MAPS_TOP_N} concorrentes para '{category}' ...")
    maps_items = find_competitors_in_maps(
        category=category, latitude=lat, longitude=lng,
        escopo=briefing.get("escopo", ""), cidade=cidade_principal,
        top_n=MAPS_TOP_N,
    )
    competitors = consolidate_maps_competitors(maps_items, briefing["url"])
    print(f"[MAPS/3b] {len(competitors)} concorrentes unicos")

    competitors_full = analyze_all_competitors(competitors, briefing["url"], briefing,
                                                 workers=5)
    return {"maps_lookup": maps_lookup, "competitors_full": competitors_full}


# ----------------------------------------------------------------------
# Modo 3: Google Ads (posts patrocinados via scraping do navegador)
# ----------------------------------------------------------------------
def consolidate_ads_competitors(ads_results: list, client_url: str) -> list:
    """Converte a saida do research_ads_for_keywords em formato compativel
    com analyze_all_competitors. Deduplica por dominio.
    `ads_results` = [{keyword, ads: [{title, domain, url, ...}], ...}, ...]"""
    cli = normalize_domain(client_url)
    by_dom = {}
    for entry in ads_results:
        kw = entry.get("keyword", "")
        for rank, ad in enumerate(entry.get("ads", []), 1):
            dom = (ad.get("domain") or "").lower().lstrip(".")
            if dom.startswith("www."):
                dom = dom[4:]
            if not dom or dom == cli or dom in ADS_DOMAINS_IGNORE:
                continue
            if dom not in by_dom:
                by_dom[dom] = {
                    "domain": dom,
                    "url": ad.get("url", "") or f"https://{dom}",
                    "title": ad.get("title", ""),
                    "keywords": [],
                    "best_rank": rank,
                    "ads_titles": [],
                    "displayed_urls": set(),
                }
            by_dom[dom]["keywords"].append({"kw": kw, "rank": rank})
            by_dom[dom]["ads_titles"].append(ad.get("title", ""))
            disp = ad.get("displayed_url", "")
            if disp:
                by_dom[dom]["displayed_urls"].add(disp)
            if rank < by_dom[dom]["best_rank"]:
                by_dom[dom]["best_rank"] = rank
                by_dom[dom]["title"] = ad.get("title", "")
                if ad.get("url"):
                    by_dom[dom]["url"] = ad["url"]

    out = []
    for c in by_dom.values():
        c["displayed_urls"] = sorted(c["displayed_urls"])
        c["ads_count"] = len(c["ads_titles"])
        out.append(c)
    out.sort(key=lambda x: (len(x["keywords"]), -x["best_rank"]), reverse=True)
    return out


def run_ads_mode(briefing: dict, selected_kws: list) -> dict:
    """Pesquisa anuncios patrocinados do Google Ads para as primeiras
    MAX_KEYWORDS_ADS keywords do modo SEO (top volume).
    Retorna {keywords, competitors_full, raw_results}."""
    print("\n" + "=" * 60)
    print(" MODO 3: GOOGLE ADS (posts patrocinados via navegador)")
    print("=" * 60)

    if not ADS_AVAILABLE:
        print("[ADS] Modulo scrape_ads_browser nao encontrado. Modo Ads ignorado.")
        return {"keywords": [], "competitors_full": [], "raw_results": []}

    if not selected_kws:
        print("[ADS] Sem keywords selecionadas. Modo Ads cancelado.")
        return {"keywords": [], "competitors_full": [], "raw_results": []}

    # limita as top N keywords (por volume) para nao alongar o processo
    selected_for_ads = selected_kws[:MAX_KEYWORDS_ADS]
    kw_list = [k["keyword"] for k in selected_for_ads]
    skipped = len(selected_kws) - len(selected_for_ads)
    if skipped > 0:
        print(f"[ADS/1] Limite de {MAX_KEYWORDS_ADS} keywords aplicado "
              f"({skipped} restantes nao serao pesquisadas no Ads):")
        for k in selected_kws[MAX_KEYWORDS_ADS:]:
            print(f"        - (skip) '{k['keyword']}'")
    print(f"[ADS/1] Buscando ads para {len(kw_list)} keyword(s) (ate 2x cada):")
    for kw in kw_list:
        print(f"        - '{kw}'")

    t0 = time.time()
    raw_results = research_ads_for_keywords(kw_list)
    elapsed = time.time() - t0
    total_ads = sum(len(r.get("ads", [])) for r in raw_results)
    print(f"[ADS/1] {total_ads} ad(s) totais em {elapsed:.1f}s")

    competitors = consolidate_ads_competitors(raw_results, briefing["url"])
    print(f"[ADS/2] {len(competitors)} anunciante(s) unico(s):")
    for c in competitors:
        print(f"        - {c['domain']:<30}  ({c['ads_count']} criativo(s), "
              f"em {len(c['keywords'])} keyword(s))")

    if not competitors:
        return {"keywords": kw_list, "competitors_full": [], "raw_results": raw_results}

    competitors_full = analyze_all_competitors(competitors, briefing["url"], briefing,
                                                 workers=5)
    return {"keywords": kw_list, "competitors_full": competitors_full,
            "raw_results": raw_results}


# ----------------------------------------------------------------------
# Modo 4: LLM (Gemini sugere concorrentes diretos com base no briefing)
# ----------------------------------------------------------------------

# limite de concorrentes a pedir ao LLM
LLM_MAX_COMPETITORS = 12

# dominios genericos / plataformas que nao queremos como "concorrente"
LLM_DOMAINS_IGNORE = {
    "google.com", "google.com.br", "facebook.com", "instagram.com",
    "youtube.com", "linkedin.com", "wikipedia.org", "reclameaqui.com.br",
    "twitter.com", "x.com", "tiktok.com", "pinterest.com",
    "whatsapp.com", "telegram.org",
}


def find_competitors_via_llm(client_url: str, briefing: dict,
                              max_competitors: int = LLM_MAX_COMPETITORS) -> list:
    """Pede ao Gemini (via OpenRouter) uma lista de concorrentes diretos da
    empresa-cliente, baseado no briefing. Retorna apenas URLs validas
    (https://dominio[/path]), sem o cliente e sem dominios genericos.

    Retorno: lista de dicts no formato esperado por analyze_all_competitors:
      [{"domain", "url", "title", "keywords": [], "best_rank": 1}, ...]
    """
    print(f"[LLM/1] Pedindo concorrentes diretos ao Gemini (max {max_competitors}) ...")

    cli_norm = normalize_domain(client_url)
    cli_brand = brand_token_from_domain(client_url)

    contexto = (
        f"URL do cliente: {client_url}\n"
        f"Dominio do cliente: {cli_norm}\n"
        f"Nicho: {briefing.get('nicho', '')}\n"
        f"Nicho principal: {briefing.get('nicho_principal', '')}\n"
        f"Publico alvo: {briefing.get('publico_alvo', '')}\n"
        f"B2B/B2C: {briefing.get('b2b_b2c', '')}\n"
        f"Escopo: {briefing.get('escopo', '')}\n"
        f"Tipo: {briefing.get('tipo_negocio', '')}\n"
        f"Dores que resolve: {briefing.get('dores', '')}\n"
    )
    cidades = briefing.get("cidades") or []
    if cidades:
        contexto += f"Cidade(s) alvo: {', '.join(cidades)}\n"

    prompt = f"""Voce e um analista de mercado especializado em SEO/SEM.

Liste ATE {max_competitors} CONCORRENTES DIRETOS da empresa abaixo (mesmo nicho,
mesmo publico, mesmo tipo de produto/servico). NAO inclua plataformas genericas
(YouTube, Instagram, Wikipedia, FAQ Whatsapp, agregadores, gov.br, marketplaces
genericos, blogs de comparativos).

EMPRESA-CLIENTE:
{contexto}

REGRAS:
- Retorne SOMENTE concorrentes diretos brasileiros (ou globais ativos no Brasil).
- NAO inclua o proprio cliente ("{cli_norm}").
- NAO inclua dominios genericos: google.com, instagram.com, facebook.com,
  youtube.com, linkedin.com, wikipedia.org, reclameaqui.com.br, whatsapp.com,
  telegram.org, gov.br, capterra.com.br, etc.
- Cada URL deve ser a HOMEPAGE do concorrente (ex: https://dominio.com.br),
  nao paginas de blog/artigo.
- Use http:// ou https://.

RETORNE APENAS JSON valido no formato exato:
{{
  "concorrentes": [
    "https://dominio1.com.br",
    "https://dominio2.com.br"
  ]
}}

Sem texto antes ou depois.
"""

    try:
        raw = call_openrouter(prompt, max_tokens=1500, temperature=0.3)
        data = extract_json(raw)
    except Exception as e:
        print(f"[LLM/1] Falha ao chamar/parsear LLM: {e}")
        return []

    urls = data.get("concorrentes") or []
    if isinstance(urls, str):
        urls = [u.strip() for u in re.split(r"[,\n]+", urls) if u.strip()]

    out = []
    seen = set()
    for u in urls:
        u = (u or "").strip()
        if not u:
            continue
        if not u.startswith(("http://", "https://")):
            u = "https://" + u.lstrip("/")
        try:
            parsed = urlparse(u)
        except Exception:
            continue
        dom = normalize_domain(parsed.netloc)
        if not dom:
            continue
        if dom == cli_norm:
            continue
        if dom in LLM_DOMAINS_IGNORE:
            continue
        # heuristica extra: nao trazer o brand do cliente em subdomínios
        if cli_brand and cli_brand in dom.replace(".", ""):
            continue
        if dom in seen:
            continue
        seen.add(dom)
        out.append({
            "domain": dom,
            "url": f"https://{dom}",
            "title": "",
            "keywords": [],
            "best_rank": 1,
        })
    print(f"[LLM/2] {len(out)} concorrente(s) validos retornados pelo LLM:")
    for c in out:
        print(f"        - {c['domain']}")
    return out


def run_llm_mode(briefing: dict) -> dict:
    """Modo 4: pede ao Gemini uma lista de concorrentes diretos da empresa."""
    print("\n" + "=" * 60)
    print(" MODO 4: LLM (Gemini sugere concorrentes diretos)")
    print("=" * 60)

    client_url = briefing.get("url", "")
    if not client_url:
        print("[LLM] URL do cliente ausente. Modo LLM cancelado.")
        return {"competitors_full": []}

    competitors = find_competitors_via_llm(client_url, briefing)
    if not competitors:
        print("[LLM] Nenhum concorrente sugerido. Modo LLM cancelado.")
        return {"competitors_full": []}

    competitors_full = analyze_all_competitors(competitors, client_url, briefing,
                                                workers=5)
    return {"competitors_full": competitors_full}


# ----------------------------------------------------------------------
# Merge entre os quatro modos
# ----------------------------------------------------------------------
_FONTE_ORDER = ["SEO", "Maps", "Ads", "LLM"]


def _merge_one(by_dom: dict, comp: dict, fonte_label: str):
    """Helper: incorpora 1 concorrente, atualizando fonte combinada."""
    d = comp["domain"]
    if d in by_dom:
        existing = by_dom[d]
        merged_kws = list(existing.get("keywords", [])) + list(comp.get("keywords", []))
        existing["keywords"] = merged_kws
        # combina fonte (set ordenado: SEO, Maps, Ads, LLM)
        parts = set(existing.get("fonte", "").split(" + ")) if existing.get("fonte") else set()
        parts.add(fonte_label)
        parts.discard("")
        existing["fonte"] = " + ".join([p for p in _FONTE_ORDER if p in parts])
        # campos extras de outras fontes
        for f in ("maps_rank", "category", "rating", "rating_count",
                  "address", "phone", "place_id",
                  "ads_count", "ads_titles", "displayed_urls"):
            if comp.get(f) and not existing.get(f):
                existing[f] = comp[f]
    else:
        by_dom[d] = {**comp, "fonte": fonte_label}


def merge_competitors(seo_list: list, maps_list: list,
                       ads_list: list = None, llm_list: list = None) -> list:
    """Une as N listas deduplicando por dominio. Cada item ganha campo
    'fonte' = combinacao ordenada de {SEO, Maps, Ads, LLM}."""
    by_dom = {}
    for c in seo_list:
        _merge_one(by_dom, c, "SEO")
    for c in maps_list:
        _merge_one(by_dom, c, "Maps")
    for c in (ads_list or []):
        _merge_one(by_dom, c, "Ads")
    for c in (llm_list or []):
        _merge_one(by_dom, c, "LLM")

    out = list(by_dom.values())
    # ordenar: Alto > Medio > Baixo, depois quantidade de fontes, depois nº keywords
    sim_rank = {"Alto": 3, "Medio": 2, "Baixo": 1, "—": 0}
    out.sort(key=lambda x: (
        sim_rank.get(x.get("similaridade", "—"), 0),
        len((x.get("fonte") or "").split(" + ")),
        len(x.get("keywords", [])),
    ), reverse=True)
    return out


# ----------------------------------------------------------------------
# XLSX
# ----------------------------------------------------------------------
def _write_competitors_sheet(wb: Workbook, sheet_name: str, competitors: list,
                              extra_cols: list = None):
    """Escreve uma aba de concorrentes (estrutura comum entre SEO e Maps).
    `extra_cols` permite adicionar colunas opcionais como "Fonte"."""
    ws = wb.create_sheet(sheet_name)
    base_headers = ["#", "Dominio", "Similaridade", "Perfil"]
    extras = extra_cols or []
    tail_headers = ["Nicho (LLM)", "Apareceu em / Pos", "Melhor Posicao",
                    "URL", "Titulo", "Motivo"]
    headers = base_headers + extras + tail_headers
    ws.append(headers)
    for c in ws[1]:
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    extras_lower = [e.lower() for e in extras]

    for i, comp in enumerate(competitors, 1):
        kws_str = ", ".join(f"{k['kw']} (#{k['rank']})" for k in comp.get("keywords", []))
        row = [i, comp["domain"], comp.get("similaridade", ""), comp.get("perfil", "")]
        for e in extras_lower:
            if e == "fonte":
                row.append(comp.get("fonte", ""))
            elif e == "categoria maps":
                row.append(comp.get("category", ""))
            elif e == "rating":
                rt = comp.get("rating")
                cnt = comp.get("rating_count")
                row.append(f"{rt} ({cnt})" if rt and cnt else (str(rt) if rt else ""))
            elif e == "pos maps":
                row.append(comp.get("maps_rank", ""))
            elif e == "# ads":
                row.append(comp.get("ads_count", ""))
            elif e == "titulos ads":
                row.append("\n".join((comp.get("ads_titles") or [])[:5]))
            else:
                row.append("")
        row.extend([
            comp.get("nicho", ""),
            kws_str,
            comp.get("best_rank", ""),
            comp.get("url", ""),
            comp.get("title", ""),
            comp.get("motivo", ""),
        ])
        ws.append(row)
        last = ws.max_row
        f = FILL_SIM.get(comp.get("similaridade", ""))
        if f:
            ws.cell(row=last, column=3).fill = f
        fp = FILL_PERFIL.get(comp.get("perfil", ""))
        if fp:
            ws.cell(row=last, column=4).fill = fp
        for cell in ws[last]:
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    base_w = [4, 28, 13, 13]
    extras_w = []
    for e in extras_lower:
        if e == "fonte": extras_w.append(15)
        elif e == "categoria maps": extras_w.append(22)
        elif e == "rating": extras_w.append(13)
        elif e == "pos maps": extras_w.append(9)
        elif e == "# ads": extras_w.append(8)
        elif e == "titulos ads": extras_w.append(50)
        else: extras_w.append(14)
    tail_w = [26, 52, 12, 46, 32, 56]
    widths = base_w + extras_w + tail_w
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w


def _write_traffic_sheet(wb: Workbook, traffic: dict, sheet_name: str = "Trafego SEO Marca"):
    if not traffic or not traffic.get("rows"):
        return
    ws_t = wb.create_sheet(sheet_name)
    cur = traffic.get("current_date") or "Atual"
    y1 = traffic.get("y1_date") or "12m antes"
    y2 = traffic.get("y2_date") or "24m antes"
    y3 = traffic.get("y3_date") or "36m antes"
    ws_t.append([f"Trafego organico (Semrush): cliente + concorrentes Alto unicos"])
    ws_t.merge_cells("A1:V1")
    ws_t["A1"].font = Font(bold=True, size=12, color="0F172A")
    ws_t.append([])
    ws_t.append([
        "Tipo", "Dominio", "Perfil",
        f"SEO {cur}", f"SEO {y1}", f"SEO {y2}", f"SEO {y3}",
        f"Marca {cur}", f"Marca {y1}", f"Marca {y2}", f"Marca {y3}",
        f"Total {cur}", f"Total {y1}", f"Total {y2}", f"Total {y3}",
        "Cresc SEO 1a (%)", "Cresc SEO 2a (%)", "Cresc SEO 3a (%)",
        "Cresc Marca 1a (%)", "Cresc Marca 2a (%)", "Cresc Marca 3a (%)",
    ])
    for c in ws_t[3]:
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    def _pct(now, then):
        if now is None or then is None:
            return None
        try:
            if not then or float(then) == 0:
                return None
            return round((float(now) - float(then)) / float(then) * 100, 1)
        except (TypeError, ValueError):
            return None

    client_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    for r in traffic["rows"]:
        cur_p = r.get("current") or {}
        y1_p = r.get("y1") or {}
        y2_p = r.get("y2") or {}
        y3_p = r.get("y3") or {}
        b_cur = cur_p.get("branded")
        b_y1 = y1_p.get("branded")
        b_y2 = y2_p.get("branded")
        b_y3 = y3_p.get("branded")
        g_brand_1y = r.get("growth_brand_1y", _pct(b_cur, b_y1))
        g_brand_2y = r.get("growth_brand_2y", _pct(b_cur, b_y2))
        g_brand_3y = r.get("growth_brand_3y", _pct(b_cur, b_y3))
        ws_t.append([
            r["role"].capitalize(),
            r["domain"],
            r.get("perfil", "—"),
            cur_p.get("non_branded"), y1_p.get("non_branded"), y2_p.get("non_branded"), y3_p.get("non_branded"),
            b_cur, b_y1, b_y2, b_y3,
            cur_p.get("total"), y1_p.get("total"), y2_p.get("total"), y3_p.get("total"),
            r.get("growth_seo_1y"), r.get("growth_seo_2y"), r.get("growth_seo_3y"),
            g_brand_1y, g_brand_2y, g_brand_3y,
        ])
        last = ws_t.max_row
        for c in ws_t[last]:
            c.border = BORDER
            c.alignment = Alignment(vertical="center")
        if r["role"] == "cliente":
            for c in ws_t[last]:
                c.fill = client_fill
                c.font = Font(bold=True)
        else:
            fp = FILL_PERFIL.get(r.get("perfil", ""))
            if fp:
                ws_t.cell(row=last, column=3).fill = fp
        for col in range(4, 16):
            ws_t.cell(row=last, column=col).number_format = "#,##0"
        for col in (16, 17, 18, 19, 20, 21):
            v = ws_t.cell(row=last, column=col).value
            ws_t.cell(row=last, column=col).number_format = "+0.0\\%;-0.0\\%;0"
            if isinstance(v, (int, float)):
                color = "16A34A" if v >= 0 else "DC2626"
                ws_t.cell(row=last, column=col).font = Font(
                    bold=(r["role"] == "cliente"), color=color)

    ws_t.column_dimensions["A"].width = 13
    ws_t.column_dimensions["B"].width = 30
    ws_t.column_dimensions["C"].width = 14
    for col_letter in list("DEFGHIJKLMNO"):
        ws_t.column_dimensions[col_letter].width = 13
    for col_letter in ("P", "Q", "R", "S", "T", "U", "V"):
        ws_t.column_dimensions[col_letter].width = 16


def _write_growth_sheet(wb: Workbook, traffic: dict, sheet_name: str = "Crescimento"):
    if not traffic or not traffic.get("rows"):
        return
    ws = wb.create_sheet(sheet_name)
    ws.append(["Crescimento SEO e Marca"])
    ws.merge_cells("A1:K1")
    ws["A1"].font = Font(bold=True, size=12, color="0F172A")
    ws.append([])
    ws.append([
        "Tipo", "Dominio", "Perfil", "Idade",
        "SEO Atual", "Marca Atual",
        "Cresc SEO 1a (%)", "Cresc SEO 2a (%)", "Cresc SEO 3a (%)",
        "Cresc Marca 1a (%)", "Cresc Marca 2a (%)", "Cresc Marca 3a (%)",
    ])
    for c in ws[3]:
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    client_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    for r in traffic["rows"]:
        cur_p = r.get("current") or {}
        y1_p = r.get("y1") or {}
        y2_p = r.get("y2") or {}
        y3_p = r.get("y3") or {}
        b_cur = cur_p.get("branded") or 0
        b_y1 = y1_p.get("branded") or 0
        b_y2 = y2_p.get("branded") or 0
        b_y3 = y3_p.get("branded") or 0
        if b_y3:
            idade = "3+"
        elif b_y2:
            idade = "2"
        elif b_y1:
            idade = "1"
        else:
            idade = "0"
        ws.append([
            r["role"].capitalize(),
            r["domain"],
            r.get("perfil", "—"),
            idade,
            cur_p.get("non_branded"),
            b_cur,
            r.get("growth_seo_1y"),
            r.get("growth_seo_2y"),
            r.get("growth_seo_3y"),
            r.get("growth_brand_1y"),
            r.get("growth_brand_2y"),
            r.get("growth_brand_3y"),
        ])
        last = ws.max_row
        for c in ws[last]:
            c.border = BORDER
            c.alignment = Alignment(vertical="center")
        if r["role"] == "cliente":
            for c in ws[last]:
                c.fill = client_fill
                c.font = Font(bold=True)
        else:
            fp = FILL_PERFIL.get(r.get("perfil", ""))
            if fp:
                ws.cell(row=last, column=3).fill = fp
        for col in (5, 6):
            ws.cell(row=last, column=col).number_format = "#,##0"
        for col in (7, 8, 9, 10, 11, 12):
            v = ws.cell(row=last, column=col).value
            ws.cell(row=last, column=col).number_format = "+0.0\\%;-0.0\\%;0"
            if isinstance(v, (int, float)):
                color = "16A34A" if v >= 0 else "DC2626"
                ws.cell(row=last, column=col).font = Font(
                    bold=(r["role"] == "cliente"), color=color)

    for col_letter, width in {
        "A": 13, "B": 30, "C": 14, "D": 8, "E": 13, "F": 13,
        "G": 16, "H": 16, "I": 16, "J": 17, "K": 17, "L": 17,
    }.items():
        ws.column_dimensions[col_letter].width = width


def write_unified_xlsx(out_path: Path, client_input: str, briefing: dict,
                        seo_result: dict, maps_result: dict, ads_result: dict,
                        llm_result: dict,
                        merged: list, traffic: dict, client_profile: str,
                        mode_key: str = "completo", mode_cfg: dict | None = None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"
    mode_cfg = mode_cfg or EXECUTION_MODES["completo"]
    channels_str = " + ".join(mode_cfg.get("channels") or [])

    ws.append([f"Concorrentes Finder - {mode_cfg.get('label', mode_key).upper()} ({channels_str})"])
    ws.merge_cells("A1:B1")
    ws["A1"].font = Font(bold=True, size=14, color="0F172A")
    ws.append([])
    ws.append(["Cliente (input)", client_input])
    ws.append(["URL cliente", briefing.get("url", "")])
    ws.append(["Nicho principal", briefing.get("nicho_principal", "")])
    ws.append(["Escopo", briefing.get("escopo", "")])
    ws.append(["Cidade(s)", ", ".join(briefing.get("cidades") or [])])
    ws.append(["Perfil do cliente", client_profile])
    ws.append(["Data", datetime.now().strftime("%d/%m/%Y %H:%M")])
    ws.append([])

    # SEO mode
    ws.append([f"MODO 1 - SEO Organico (SERP top {mode_cfg.get('seo_depth', SERP_DEPTH_TOP10)})"])
    ws["A10"].font = Font(bold=True, color="0F766E")
    selected = seo_result.get("selected_kws") or []
    seo_comp = seo_result.get("competitors_full") or []
    sim_seo = {"Alto": 0, "Medio": 0, "Baixo": 0}
    for c in seo_comp:
        s = c.get("similaridade", "—")
        if s in sim_seo:
            sim_seo[s] += 1
    ws.append(["Keywords usadas", len(selected)])
    for k in selected:
        ws.append([f"   - {k['keyword']}", f"vol {k.get('volume')}"])
    ws.append(["Concorrentes encontrados (SEO)", len(seo_comp)])
    ws.append(["SEO - Alto", sim_seo["Alto"]])
    ws.append(["SEO - Medio", sim_seo["Medio"]])
    ws.append(["SEO - Baixo", sim_seo["Baixo"]])
    ws.append([])

    # Maps mode
    ws.append(["MODO 2 - Google Maps (categoria + raio)"])
    last_row = ws.max_row
    ws.cell(row=last_row, column=1).font = Font(bold=True, color="0F766E")
    maps_lookup = maps_result.get("maps_lookup") or {}
    maps_comp = maps_result.get("competitors_full") or []
    if maps_lookup:
        ws.append(["Maps - Empresa", maps_lookup.get("title", "")])
        ws.append(["Maps - Categoria", maps_lookup.get("category", "")])
        extras = maps_lookup.get("additional_categories") or []
        ws.append(["Maps - Cat. adicionais", ", ".join(extras) if extras else ""])
        ws.append(["Maps - Endereco", maps_lookup.get("address", "")])
        ws.append(["Maps - Latitude", maps_lookup.get("latitude")])
        ws.append(["Maps - Longitude", maps_lookup.get("longitude")])
        rt = (maps_lookup.get("rating") or {}).get("value")
        rt_c = (maps_lookup.get("rating") or {}).get("votes_count")
        ws.append(["Maps - Rating", f"{rt} ({rt_c} avaliacoes)" if rt else ""])
    sim_maps = {"Alto": 0, "Medio": 0, "Baixo": 0}
    for c in maps_comp:
        s = c.get("similaridade", "—")
        if s in sim_maps:
            sim_maps[s] += 1
    ws.append(["Concorrentes encontrados (Maps)", len(maps_comp)])
    ws.append(["Maps - Alto", sim_maps["Alto"]])
    ws.append(["Maps - Medio", sim_maps["Medio"]])
    ws.append(["Maps - Baixo", sim_maps["Baixo"]])
    ws.append([])

    # Ads mode
    ws.append(["MODO 3 - Google Ads (posts patrocinados)"])
    last_row = ws.max_row
    ws.cell(row=last_row, column=1).font = Font(bold=True, color="0F766E")
    ads_comp = ads_result.get("competitors_full") or []
    ads_raw = ads_result.get("raw_results") or []
    total_creatives = sum(len(r.get("ads", [])) for r in ads_raw)
    ws.append(["Keywords pesquisadas", len(ads_result.get("keywords") or [])])
    ws.append(["Criativos totais", total_creatives])
    sim_ads = {"Alto": 0, "Medio": 0, "Baixo": 0}
    for c in ads_comp:
        s = c.get("similaridade", "—")
        if s in sim_ads:
            sim_ads[s] += 1
    ws.append(["Anunciantes unicos", len(ads_comp)])
    ws.append(["Ads - Alto", sim_ads["Alto"]])
    ws.append(["Ads - Medio", sim_ads["Medio"]])
    ws.append(["Ads - Baixo", sim_ads["Baixo"]])
    ws.append([])

    # LLM mode
    ws.append(["MODO 4 - LLM (Gemini sugere concorrentes diretos)"])
    last_row = ws.max_row
    ws.cell(row=last_row, column=1).font = Font(bold=True, color="0F766E")
    llm_comp = llm_result.get("competitors_full") or []
    sim_llm = {"Alto": 0, "Medio": 0, "Baixo": 0}
    for c in llm_comp:
        s = c.get("similaridade", "—")
        if s in sim_llm:
            sim_llm[s] += 1
    ws.append(["Concorrentes sugeridos pelo LLM", len(llm_comp)])
    ws.append(["LLM - Alto", sim_llm["Alto"]])
    ws.append(["LLM - Medio", sim_llm["Medio"]])
    ws.append(["LLM - Baixo", sim_llm["Baixo"]])
    ws.append([])

    # Unificado
    ws.append(["UNIFICADO"])
    last_row = ws.max_row
    ws.cell(row=last_row, column=1).font = Font(bold=True, color="0F766E")
    fonte_counts = {}
    for c in merged:
        f = c.get("fonte", "SEO")
        fonte_counts[f] = fonte_counts.get(f, 0) + 1
    ws.append(["Total unico", len(merged)])
    # imprime cada combinacao ordenada por count
    for f, n in sorted(fonte_counts.items(), key=lambda kv: -kv[1]):
        ws.append([f"  {f}", n])
    sim_all = {"Alto": 0, "Medio": 0, "Baixo": 0}
    for c in merged:
        s = c.get("similaridade", "—")
        if s in sim_all:
            sim_all[s] += 1
    ws.append(["Unificado - Alto", sim_all["Alto"]])
    ws.append(["Unificado - Medio", sim_all["Medio"]])
    ws.append(["Unificado - Baixo", sim_all["Baixo"]])

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 60

    # Abas: SEO, Maps, Ads, Unificado
    if seo_comp:
        _write_competitors_sheet(wb, "Concorrentes SEO", seo_comp)
    if maps_comp:
        _write_competitors_sheet(wb, "Concorrentes Maps", maps_comp,
                                  extra_cols=["Pos Maps", "Categoria Maps", "Rating"])
    if ads_comp:
        _write_competitors_sheet(wb, "Concorrentes Ads", ads_comp,
                                  extra_cols=["# Ads", "Titulos Ads"])
    if llm_comp:
        _write_competitors_sheet(wb, "Concorrentes LLM", llm_comp)
    if merged:
        _write_competitors_sheet(wb, "Concorrentes Unificado", merged,
                                  extra_cols=["Fonte", "Categoria Maps", "Rating", "# Ads"])

    # Trafego (cliente + Alto unicos do unificado)
    _write_traffic_sheet(wb, traffic)
    _write_growth_sheet(wb, traffic)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    saved = out_path
    idx = 1
    while True:
        try:
            wb.save(saved)
            break
        except PermissionError:
            saved = out_path.with_name(f"{out_path.stem}_{idx}{out_path.suffix}")
            idx += 1
    return saved


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    if len(sys.argv) < 2:
        print("Uso: python find_concorrentes_all.py <nome_cliente> [modo]")
        print("Ex:  python find_concorrentes_all.py cnempreendimentos")
        print("Ex:  python find_concorrentes_all.py chatguru nacional")
        sys.exit(1)

    client_input = sys.argv[1].strip()
    mode_key = (sys.argv[2].strip().lower() if len(sys.argv) >= 3 else "completo")
    if mode_key not in EXECUTION_MODES:
        print(f"ERRO: modo '{mode_key}' invalido. Opcoes: {', '.join(EXECUTION_MODES)}")
        sys.exit(1)
    mode_cfg = EXECUTION_MODES[mode_key]

    print(f"\n=== Concorrentes Finder: '{client_input}' | modo {mode_cfg['label']} ===\n")
    t0 = time.time()

    # 1) localiza e le briefing
    xlsx_path = find_client_xlsx(client_input)
    briefing = read_briefing_from_xlsx(xlsx_path)
    if not briefing.get("url"):
        print("ERRO: URL do cliente nao encontrada no briefing.")
        sys.exit(1)
    print(f"[1] Briefing: nicho='{briefing.get('nicho_principal','')}', "
          f"escopo={briefing.get('escopo')}, cidades={briefing.get('cidades')}")
    client_slug = xlsx_path.parent.name

    # 2) modo SEO
    channels = set(mode_cfg.get("channels") or [])

    seo_result = run_seo_mode(briefing, xlsx_path, serp_depth=mode_cfg.get("seo_depth", SERP_DEPTH_TOP10)) if "SEO" in channels else {"selected_kws": [], "competitors_full": []}

    # 3) modo Maps
    maps_result = run_maps_mode(briefing, client_slug) if "Maps" in channels else {"maps_lookup": {}, "competitors_full": []}

    # 4) modo Ads (usa as MESMAS keywords que o SEO usou)
    ads_result = run_ads_mode(briefing, seo_result.get("selected_kws") or []) if "Ads" in channels else {"keywords": [], "competitors_full": [], "raw_results": []}

    # 5) modo LLM (Gemini sugere concorrentes diretos)
    llm_result = run_llm_mode(briefing) if "LLM" in channels else {"competitors_full": []}

    # 6) merge
    print("\n" + "=" * 60)
    print(f" MERGE: {' + '.join(mode_cfg.get('channels') or [])}")
    print("=" * 60)
    merged = merge_competitors(seo_result["competitors_full"],
                                maps_result["competitors_full"],
                                ads_result["competitors_full"],
                                llm_result.get("competitors_full") or [])
    print(f"[MERGE] {len(merged)} concorrentes unicos:")
    fonte_counts = {}
    for c in merged:
        f = c.get("fonte", "SEO")
        fonte_counts[f] = fonte_counts.get(f, 0) + 1
    for k, v in sorted(fonte_counts.items(), key=lambda kv: -kv[1]):
        print(f"        {k}: {v}")
    print(f"[MERGE] Concorrentes Alto unicos:")
    for c in [c for c in merged if c.get("similaridade") == "Alto"]:
        print(f"          {c['fonte']:<18}  {c['domain']:<30}  ({c.get('perfil','')})")

    # 6) perfil do cliente
    print("\n[6] Classificando perfil do cliente ...")
    client_url = briefing["url"]
    if not client_url.startswith(("http://", "https://")):
        client_url = "https://" + client_url
    try:
        site = scrape_site(client_url)
        client_profile = classify_client_profile(client_url, site, briefing)
    except Exception as e:
        print(f"[6] Falha ao classificar cliente: {e}")
        client_profile = "—"
    print(f"[6] Cliente: {client_profile}")

    # 7) trafego Semrush (cliente + Alto unicos do merged)
    print("\n[7] Trafego Semrush 3 periodos (cliente + Alto unicos):")
    traffic = build_3period_traffic(client_url, merged, client_profile,
                                     max_competitors=mode_cfg.get("traffic_limit", 15))

    # 8) saida
    suffix = "" if mode_key == "completo" else f"-{mode_key}"
    out_path = OUT_DIR / client_slug / f"concorrentes-all-{client_slug}{suffix}.xlsx"
    print(f"\n[OUT] Escrevendo: {out_path}")
    saved = write_unified_xlsx(out_path, client_input, briefing,
                                 seo_result, maps_result, ads_result, llm_result,
                                 merged, traffic, client_profile,
                                 mode_key=mode_key, mode_cfg=mode_cfg)
    print(f"[OUT] XLSX salvo: {saved}")
    try:
        from gerar_html_concorrentes import generate_html_from_xlsx
        html_path = generate_html_from_xlsx(saved)
        print(f"[OUT] HTML salvo: {html_path}")
    except Exception as e:
        print(f"[OUT] HTML nao gerado: {e}")
        html_path = None

    print(f"\n=== Concluido em {time.time()-t0:.1f}s ===")
    print(f"XLSX: {saved}")
    if html_path:
        print(f"HTML: {html_path}")


if __name__ == "__main__":
    main()
