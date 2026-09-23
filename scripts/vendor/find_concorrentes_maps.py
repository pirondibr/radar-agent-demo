# -*- coding: utf-8 -*-
"""
Concorrentes Finder - modo Google Maps

Etapas:
  1. Localiza XLSX do cliente (mesma logica do find_concorrentes.py).
  2. Le briefing para extrair nome do cliente, URL e cidade(s) alvo.
  3. Google Maps SERP (DataForSEO) busca o cliente especifico para descobrir:
       - categoria oficial do estabelecimento (ex: "Construtora")
       - latitude/longitude do cliente
  4. Nova busca no Maps usando a CATEGORIA + coordenadas do cliente (raio ~15km
     para local; sem coord para nacional). Pega top 10 com site.
  5. Roda similaridade + perfil via LLM para cada concorrente.
  6. Busca tragego Semrush 3 periodos (atual, 12m, 24m) para cliente + Alto.
  7. Salva XLSX com 3 abas.

Uso:
    python find_concorrentes_maps.py <nome_cliente>
    python find_concorrentes_maps.py cnempreendimentos
    python find_concorrentes_maps.py apetsaude
"""

import sys
import re
import time
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse
from difflib import SequenceMatcher

import requests
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ----------------------------------------------------------------------
# Reuso do seo_pipeline
# ----------------------------------------------------------------------
_THIS_DIR = Path(__file__).resolve().parent
from workspace_paths import setup_workspace  # noqa: E402

setup_workspace()

from seo_pipeline import (  # type: ignore  # noqa: E402
    DATAFORSEO_USER, DATAFORSEO_PASS, LOCATION_BRAZIL,
    normalize_domain, analyze_similarity, classify_client_profile,
    fetch_semrush_monthly_trend, scrape_site,
)
from find_concorrentes import (  # type: ignore  # noqa: E402
    find_client_xlsx, read_briefing_from_xlsx,
    analyze_all_competitors, pick_3_periods, build_3period_traffic,
    DOMAINS_TO_IGNORE, FILL_SIM, FILL_PERFIL, HEAD_FONT, HEAD_FILL, BORDER,
)


# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
DATAFORSEO_MAPS_URL = "https://api.dataforseo.com/v3/serp/google/maps/live/advanced"
OUT_DIR = _THIS_DIR / "output"

MAPS_RADIUS_KM = 15        # raio padrao para busca local
MAPS_RADIUS_FALLBACK = 50  # raio maior se primeira busca trouxer poucos
MAPS_TOP_N = 10            # quantos concorrentes pegar


# ----------------------------------------------------------------------
# DataForSEO Maps
# ----------------------------------------------------------------------
def maps_search(keyword: str, location_coordinate: str = None,
                depth: int = 20) -> list:
    """Chamada ao endpoint Maps live/advanced. Retorna lista de items."""
    payload = [{
        "language_code": "pt",
        "location_code": LOCATION_BRAZIL,
        "keyword": keyword,
        "depth": depth,
        "device": "desktop",
    }]
    if location_coordinate:
        payload[0]["location_coordinate"] = location_coordinate
    try:
        r = requests.post(DATAFORSEO_MAPS_URL, json=payload,
                          auth=(DATAFORSEO_USER, DATAFORSEO_PASS), timeout=60)
        r.raise_for_status()
        js = r.json()
    except Exception as e:
        print(f"[maps] erro: {e}")
        return []
    if js.get("status_code") != 20000:
        print(f"[maps] status_code: {js.get('status_message')}")
        return []
    tasks = js.get("tasks") or []
    if not tasks or tasks[0].get("status_code") != 20000:
        msg = tasks[0].get("status_message") if tasks else "sem tasks"
        if "No Search Results" not in str(msg):
            print(f"[maps] task: {msg}")
        return []
    result = tasks[0].get("result") or []
    if not result:
        return []
    return result[0].get("items") or []


def find_client_in_maps(client_name: str, client_url: str,
                         briefing: dict) -> dict:
    """Encontra o cliente no Maps para extrair categoria + coordenadas.
    Estrategias (tenta em ordem):
        1. <nome_cliente> <cidade principal>
        2. <nome_cliente> + nicho
        3. <nome_cliente> sozinho
    """
    cli_dom = normalize_domain(client_url)
    cli_root = cli_dom.split(".")[0]  # ex: cnempreendimentos
    cidades = briefing.get("cidades") or []
    cidade_principal = cidades[0] if cidades else ""

    queries = []
    if cidade_principal:
        queries.append(f"{client_name} {cidade_principal}")
    queries.append(client_name)
    if briefing.get("nicho_principal"):
        queries.append(f"{client_name} {briefing['nicho_principal']}")

    seen_q = set()
    for q in queries:
        q = q.strip()
        if not q or q.lower() in seen_q:
            continue
        seen_q.add(q.lower())
        print(f"[3a]   tentando: '{q}'")
        items = maps_search(q, depth=10)
        if not items:
            continue

        # Procura o item cujo dominio bate com o do cliente, OU cujo nome eh
        # muito similar
        best = None
        for it in items:
            it_dom = normalize_domain(it.get("url", ""))
            if it_dom and cli_dom in it_dom:
                best = it
                break
            # match por nome
            t = (it.get("title") or "").lower()
            sim = SequenceMatcher(None, t, client_name.lower()).ratio()
            if sim >= 0.7 or cli_root in t.replace(" ", ""):
                if not best or sim > SequenceMatcher(
                        None, (best.get("title") or "").lower(),
                        client_name.lower()).ratio():
                    best = it
        if best:
            return best
        # nao achou match, mas se o primeiro item tem boa relevancia, retorna ele
        if len(items) == 1:
            return items[0]

    return {}


def find_competitors_in_maps(category: str, latitude: float, longitude: float,
                              escopo: str, cidade: str = "",
                              top_n: int = MAPS_TOP_N) -> list:
    """Top N concorrentes no Maps via categoria + coord (Local) ou
    categoria + cidade (Nacional)."""
    coord_str = None
    if escopo == "Local" and latitude and longitude:
        coord_str = f"{latitude},{longitude},{MAPS_RADIUS_KM}"
        keyword = category
    else:
        # Nacional: tenta sem coordenadas
        keyword = f"{category} {cidade}".strip() if cidade else category

    print(f"[3b]   busca: '{keyword}'" + (f"  coord={coord_str}" if coord_str else ""))
    items = maps_search(keyword, location_coordinate=coord_str, depth=20)

    # se vazio com 15km, tenta raio maior
    if not items and coord_str:
        coord_str = f"{latitude},{longitude},{MAPS_RADIUS_FALLBACK}"
        print(f"[3b]   ampliando raio para {MAPS_RADIUS_FALLBACK}km")
        items = maps_search(keyword, location_coordinate=coord_str, depth=20)

    # filtra: precisa ter url, e a url precisa ser um site real (nao IG/FB sozinho)
    out = []
    for it in items:
        url = (it.get("url") or "").strip()
        if not url:
            continue
        dom = normalize_domain(url)
        if dom in DOMAINS_TO_IGNORE:
            continue
        out.append(it)
    return out[:top_n]


# ----------------------------------------------------------------------
# Consolida resultados Maps em "competitors" compativel com analyze_all
# ----------------------------------------------------------------------
def consolidate_maps_competitors(maps_items: list, client_url: str) -> list:
    """Converte items do Maps para o formato esperado por analyze_all_competitors.
    Deduplica por dominio (mantem primeira aparicao)."""
    cli = normalize_domain(client_url)
    seen = set()
    out = []
    for rank, it in enumerate(maps_items, 1):
        url = (it.get("url") or "").strip()
        dom = normalize_domain(url)
        if not dom or dom == cli or dom in seen:
            continue
        seen.add(dom)
        rating = (it.get("rating") or {}).get("value")
        out.append({
            "domain": dom,
            "url": url,
            "title": it.get("title", ""),
            "category": it.get("category", ""),
            "rating": rating,
            "rating_count": (it.get("rating") or {}).get("votes_count"),
            "address": it.get("address", ""),
            "phone": it.get("phone", ""),
            "maps_rank": rank,
            "place_id": it.get("place_id", ""),
            "keywords": [{"kw": "Google Maps", "rank": rank}],
            "best_rank": rank,
        })
    return out


# ----------------------------------------------------------------------
# XLSX
# ----------------------------------------------------------------------
def write_xlsx(out_path: Path, client_input: str, briefing: dict,
               maps_lookup: dict, competitors_full: list, traffic: dict):
    wb = Workbook()

    # ---- Aba Resumo
    ws = wb.active
    ws.title = "Resumo"
    ws.append(["Concorrentes Finder - modo Google Maps"])
    ws.merge_cells("A1:B1")
    ws["A1"].font = Font(bold=True, size=14, color="0F172A")
    ws.append([])
    ws.append(["Cliente (input)", client_input])
    ws.append(["URL cliente", briefing.get("url", "")])
    ws.append(["Nicho principal", briefing.get("nicho_principal", "")])
    ws.append(["Escopo", briefing.get("escopo", "")])
    ws.append(["Cidade(s)", ", ".join(briefing.get("cidades") or [])])
    ws.append(["Data", datetime.now().strftime("%d/%m/%Y %H:%M")])
    ws.append([])
    ws.append(["Resultado da busca Maps - cliente"])
    ws.append(["Maps - Titulo", maps_lookup.get("title", "")])
    ws.append(["Maps - Categoria", maps_lookup.get("category", "")])
    cats_extra = maps_lookup.get("additional_categories") or []
    ws.append(["Maps - Categorias adicionais", ", ".join(cats_extra) if cats_extra else ""])
    ws.append(["Maps - Endereco", maps_lookup.get("address", "")])
    ws.append(["Maps - Telefone", maps_lookup.get("phone", "")])
    ws.append(["Maps - Latitude", maps_lookup.get("latitude")])
    ws.append(["Maps - Longitude", maps_lookup.get("longitude")])
    rating = (maps_lookup.get("rating") or {}).get("value")
    rating_votes = (maps_lookup.get("rating") or {}).get("votes_count")
    ws.append(["Maps - Rating", f"{rating} ({rating_votes} avaliacoes)" if rating else ""])
    ws.append([])
    ws.append(["Concorrentes encontrados"])
    ws.append(["Total unico", len(competitors_full)])
    sim_counts = {"Alto": 0, "Medio": 0, "Baixo": 0, "—": 0}
    for c in competitors_full:
        s = c.get("similaridade", "—")
        sim_counts[s] = sim_counts.get(s, 0) + 1
    ws.append(["Similaridade Alto", sim_counts.get("Alto", 0)])
    ws.append(["Similaridade Medio", sim_counts.get("Medio", 0)])
    ws.append(["Similaridade Baixo", sim_counts.get("Baixo", 0)])

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 60

    # ---- Aba Concorrentes Maps
    ws_c = wb.create_sheet("Concorrentes Maps")
    headers = ["#", "Pos Maps", "Dominio", "Empresa", "Categoria Maps",
               "Similaridade", "Perfil", "Rating", "Nicho (LLM)", "URL",
               "Endereco", "Telefone", "Motivo"]
    ws_c.append(headers)
    for c in ws_c[1]:
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    for i, comp in enumerate(competitors_full, 1):
        rating_txt = ""
        if comp.get("rating"):
            rt = comp["rating"]
            cnt = comp.get("rating_count")
            rating_txt = f"{rt}" + (f" ({cnt})" if cnt else "")
        ws_c.append([
            i, comp.get("maps_rank", ""), comp["domain"], comp.get("title", ""),
            comp.get("category", ""), comp.get("similaridade", ""),
            comp.get("perfil", ""), rating_txt, comp.get("nicho", ""),
            comp.get("url", ""), comp.get("address", ""),
            comp.get("phone", ""), comp.get("motivo", ""),
        ])
        last = ws_c.max_row
        f = FILL_SIM.get(comp.get("similaridade", ""))
        if f:
            ws_c.cell(row=last, column=6).fill = f
        fp = FILL_PERFIL.get(comp.get("perfil", ""))
        if fp:
            ws_c.cell(row=last, column=7).fill = fp
        for cell in ws_c[last]:
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    widths = [4, 9, 28, 32, 22, 14, 14, 12, 22, 44, 50, 16, 60]
    for i, w in enumerate(widths, 1):
        ws_c.column_dimensions[chr(64 + i)].width = w

    # ---- Aba Trafego SEO Marca (cliente + Alto)
    if traffic and traffic.get("rows"):
        ws_t = wb.create_sheet("Trafego SEO Marca")
        cur = traffic.get("current_date") or "Atual"
        y1 = traffic.get("y1_date") or "12m antes"
        y2 = traffic.get("y2_date") or "24m antes"
        ws_t.append(["Trafego organico (Semrush): SEO + Marca - 3 periodos - concorrentes via Google Maps"])
        ws_t.merge_cells("A1:N1")
        ws_t["A1"].font = Font(bold=True, size=12, color="0F172A")
        ws_t.append([])
        ws_t.append([
            "Tipo", "Dominio", "Perfil",
            f"SEO {cur}", f"SEO {y1}", f"SEO {y2}",
            f"Marca {cur}", f"Marca {y1}", f"Marca {y2}",
            f"Total {cur}", f"Total {y1}", f"Total {y2}",
            "Cresc SEO 1a (%)", "Cresc SEO 2a (%)",
        ])
        for c in ws_t[3]:
            c.font = HEAD_FONT
            c.fill = HEAD_FILL
            c.border = BORDER
            c.alignment = Alignment(horizontal="center", wrap_text=True)

        client_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
        for r in traffic["rows"]:
            cur_p = r.get("current") or {}
            y1_p = r.get("y1") or {}
            y2_p = r.get("y2") or {}
            ws_t.append([
                r["role"].capitalize(),
                r["domain"],
                r.get("perfil", "—"),
                cur_p.get("non_branded"), y1_p.get("non_branded"), y2_p.get("non_branded"),
                cur_p.get("branded"), y1_p.get("branded"), y2_p.get("branded"),
                cur_p.get("total"), y1_p.get("total"), y2_p.get("total"),
                r.get("growth_seo_1y"), r.get("growth_seo_2y"),
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
            for col in range(4, 13):
                ws_t.cell(row=last, column=col).number_format = "#,##0"
            for col in (13, 14):
                v = ws_t.cell(row=last, column=col).value
                ws_t.cell(row=last, column=col).number_format = "+0.0\\%;-0.0\\%;0"
                if isinstance(v, (int, float)):
                    color = "16A34A" if v >= 0 else "DC2626"
                    ws_t.cell(row=last, column=col).font = Font(
                        bold=(r["role"] == "cliente"), color=color)

        ws_t.column_dimensions["A"].width = 13
        ws_t.column_dimensions["B"].width = 30
        ws_t.column_dimensions["C"].width = 14
        for col_letter in list("DEFGHIJKL"):
            ws_t.column_dimensions[col_letter].width = 13
        ws_t.column_dimensions["M"].width = 16
        ws_t.column_dimensions["N"].width = 16

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
def derive_client_name(briefing: dict, fallback_slug: str) -> str:
    """Tenta derivar um nome 'pesquisavel' do cliente para o Maps."""
    url = briefing.get("url", "")
    host = urlparse(url).netloc or url
    root = normalize_domain(host).split(".")[0]
    # tenta dividir CamelCase ou separadores
    name = re.sub(r"[-_]", " ", root)
    if not name:
        name = fallback_slug
    return name


def main():
    if len(sys.argv) < 2:
        print("Uso: python find_concorrentes_maps.py <nome_cliente>")
        print("Ex:  python find_concorrentes_maps.py cnempreendimentos")
        sys.exit(1)

    client_input = sys.argv[1].strip()
    print(f"\n=== Concorrentes Finder (Google Maps): '{client_input}' ===\n")
    t0 = time.time()

    # 1) localiza XLSX e le briefing
    xlsx_path = find_client_xlsx(client_input)
    briefing = read_briefing_from_xlsx(xlsx_path)
    if not briefing.get("url"):
        print("ERRO: URL do cliente nao encontrada no briefing.")
        sys.exit(1)
    print(f"[1] Briefing: nicho='{briefing.get('nicho_principal','')}', "
          f"escopo={briefing.get('escopo')}, "
          f"cidades={briefing.get('cidades')}")

    client_slug = xlsx_path.parent.name
    client_name = derive_client_name(briefing, client_slug)
    print(f"[2] Nome derivado do cliente para Maps: '{client_name}'")

    # 3a) busca o cliente no Maps
    print("[3a] Buscando cliente no Google Maps ...")
    maps_lookup = find_client_in_maps(client_name, briefing["url"], briefing)
    if not maps_lookup:
        print("ERRO: cliente nao encontrado no Google Maps.")
        sys.exit(1)
    category = maps_lookup.get("category", "")
    lat = maps_lookup.get("latitude")
    lng = maps_lookup.get("longitude")
    print(f"[3a] Cliente identificado: '{maps_lookup.get('title','')}'")
    print(f"     - categoria: {category}")
    print(f"     - cat. extras: {maps_lookup.get('additional_categories') or []}")
    print(f"     - lat/lng: {lat}, {lng}")
    print(f"     - endereco: {maps_lookup.get('address','')}")
    if not category:
        print("ERRO: categoria nao identificada para o cliente.")
        sys.exit(1)

    # 3b) busca concorrentes pela categoria
    cidades = briefing.get("cidades") or []
    cidade_principal = cidades[0] if cidades else ""
    print(f"\n[3b] Buscando top {MAPS_TOP_N} concorrentes via Maps "
          f"(categoria='{category}', escopo={briefing.get('escopo')}) ...")
    maps_items = find_competitors_in_maps(
        category=category, latitude=lat, longitude=lng,
        escopo=briefing.get("escopo", ""), cidade=cidade_principal,
        top_n=MAPS_TOP_N,
    )
    print(f"[3b] {len(maps_items)} concorrentes com site retornados")

    if not maps_items:
        print("ERRO: nenhum concorrente com site encontrado no Maps.")
        sys.exit(1)

    competitors = consolidate_maps_competitors(maps_items, briefing["url"])
    print(f"[3b] {len(competitors)} concorrentes unicos por dominio")
    for c in competitors:
        print(f"        - #{c['maps_rank']:>2}  {c['domain']:<30} | {c['category']:<25} | {c['title']}")

    # 4) similaridade + perfil
    competitors_full = analyze_all_competitors(competitors, briefing["url"], briefing,
                                                 workers=5)

    # 5) perfil do cliente
    print("[5] Classificando perfil do cliente ...")
    client_url = briefing["url"]
    if not client_url.startswith(("http://", "https://")):
        client_url = "https://" + client_url
    try:
        site = scrape_site(client_url)
        client_profile = classify_client_profile(client_url, site, briefing)
    except Exception as e:
        print(f"[5] Falha ao classificar cliente: {e}")
        client_profile = "—"
    print(f"[5] Cliente classificado como: {client_profile}")

    # 6) trafego Semrush 3 periodos (cliente + Alto)
    traffic = build_3period_traffic(client_url, competitors_full, client_profile)

    # 7) saida
    out_path = OUT_DIR / client_slug / f"concorrentes-maps-{client_slug}.xlsx"
    print(f"\n[OUT] Escrevendo: {out_path}")
    saved = write_xlsx(out_path, client_input, briefing, maps_lookup,
                        competitors_full, traffic)
    print(f"[OUT] XLSX salvo: {saved}")

    print(f"\n=== Concluido em {time.time()-t0:.1f}s ===")
    print(f"XLSX: {saved}")


if __name__ == "__main__":
    main()
