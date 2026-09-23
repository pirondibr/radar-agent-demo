# -*- coding: utf-8 -*-
"""
Concorrentes Finder

Recebe o nome de um cliente (slug do XLSX do pipeline SEO), abre o relatorio
existente, escolhe ate 5 keywords (3 da aba Palavras-chave + 2 da Similaridade
Top 5 que tenham concorrente Alto, todas distintas por intencao), busca o top 10
do Google para cada via DataForSEO, consolida concorrentes unicos, classifica
similaridade e perfil via LLM, e finalmente busca trafego SEO/Marca no Semrush
em 3 periodos (atual, 12m antes, 24m antes) para cliente + concorrentes Alto.

Uso:
    python find_concorrentes.py <nome_cliente>
    python find_concorrentes.py cnempreendimentos
    python find_concorrentes.py apetsaude
"""

import sys
import re
import time
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ----------------------------------------------------------------------
# Reuso de funcoes do seo_pipeline (pasta Lista automatica vertical 1)
# ----------------------------------------------------------------------
_THIS_DIR = Path(__file__).resolve().parent
from workspace_paths import setup_workspace, PIPELINE_OUTPUT_DIR  # noqa: E402

setup_workspace()

from seo_pipeline import (  # type: ignore  # noqa: E402
    fetch_serps,
    normalize_domain,
    analyze_similarity,
    classify_client_profile,
    fetch_semrush_monthly_trend,
    keyword_signature,
    dedupe_keywords_by_intent,
    scrape_site,
    brand_token_from_domain,
    is_brand_keyword,
)


# ----------------------------------------------------------------------
# CONFIG do script
# ----------------------------------------------------------------------
OUT_DIR = _THIS_DIR / "output"

SERP_DEPTH_TOP10 = 10
MAX_KEYWORDS_KW = 3       # 3 keywords da aba "Palavras-chave"
MAX_KEYWORDS_SIM = 2      # 2 keywords "Alto" da aba "Similaridade Top 5"
MAX_TOTAL_KEYWORDS = 5

MAX_COMPETITORS_FOR_TRAFFIC = 10   # quantos concorrentes Alto entram no relatorio Semrush

DOMAINS_TO_IGNORE = {
    # agregadores/portais que aparecem em SERP mas nao sao "concorrentes" no sentido tradicional
    "google.com", "google.com.br", "facebook.com", "instagram.com",
    "youtube.com", "linkedin.com", "wikipedia.org", "reclameaqui.com.br",
    "twitter.com", "x.com", "tiktok.com", "pinterest.com",
}


# ----------------------------------------------------------------------
# Localiza o XLSX do cliente
# ----------------------------------------------------------------------
def find_client_xlsx(client_input: str) -> Path:
    """Procura o XLSX mais recente do cliente. Aceita:
    - slug puro (ex: 'cnempreendimentos')
    - nome do dominio (ex: 'cnempreendimentos.net')
    - URL completa (ex: 'https://cnempreendimentos.net/')
    """
    raw = client_input.strip().lower()
    raw = raw.replace("https://", "").replace("http://", "").rstrip("/")
    raw = raw.replace("www.", "").split("/")[0]
    slug = re.sub(r"[^a-z0-9]+", "", raw.split(".")[0])

    if not slug:
        raise ValueError(f"Nao consegui derivar slug de '{client_input}'")

    candidate_dir = PIPELINE_OUTPUT_DIR / slug
    if not candidate_dir.exists():
        # fallback: procura por qualquer pasta cujo nome contenha o slug
        matches = [p for p in PIPELINE_OUTPUT_DIR.glob("*")
                   if p.is_dir() and slug in p.name.lower()]
        if not matches:
            raise FileNotFoundError(
                f"Pasta do cliente '{slug}' nao encontrada em {PIPELINE_OUTPUT_DIR}.\n"
                f"Rode primeiro a etapa 'entender o cliente' para gerar "
                f"briefing-cliente-<slug>.xlsx"
            )
        candidate_dir = matches[0]
        slug = candidate_dir.name

    # Preferir briefing de "entender o cliente"; relatorio-seo fica como fallback.
    patterns = (
        f"briefing-cliente-{slug}*.xlsx",
        "briefing-cliente-*.xlsx",
        f"relatorio-seo-{slug}*.xlsx",
        "relatorio-seo-*.xlsx",
    )
    xlsx_files = []
    for pattern in patterns:
        xlsx_files = sorted(candidate_dir.glob(pattern),
                            key=lambda p: p.stat().st_mtime, reverse=True)
        if xlsx_files:
            break
    if not xlsx_files:
        raise FileNotFoundError(
            f"Nenhum XLSX do cliente em {candidate_dir}.\n"
            f"Rode primeiro a etapa 'entender o cliente' para gerar "
            f"briefing-cliente-<slug>.xlsx"
        )

    xlsx = xlsx_files[0]
    print(f"[1] XLSX do cliente encontrado: {xlsx.name}")
    return xlsx


# ----------------------------------------------------------------------
# Leitura do XLSX do cliente
# ----------------------------------------------------------------------
def read_briefing_from_xlsx(xlsx_path: Path) -> dict:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["Briefing"]
    fields = {}
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        key = str(row[0]).strip()
        val = "" if (len(row) < 2 or row[1] is None) else str(row[1]).strip()
        fields[key] = val

    cidades_raw = fields.get("Cidades alvo", "")
    cidades = [c.strip() for c in re.split(r"[,;]", cidades_raw) if c.strip()]

    def find_by_prefix(prefix: str) -> str:
        """Busca campo cujo label COMECA com prefix (case-insensitive),
        para sobreviver a variacoes de acento/parenteses."""
        prefix_l = prefix.lower()
        for k, v in fields.items():
            if k.lower().startswith(prefix_l):
                return v
        return ""

    briefing = {
        "url": fields.get("URL", ""),
        "nicho": find_by_prefix("Nicho") if not find_by_prefix("Nicho principal") else find_by_prefix("Nicho "),
        "nicho_principal": find_by_prefix("Nicho principal"),
        "nicho_secundario": find_by_prefix("Nicho secundario"),
        "tipo_negocio": find_by_prefix("Servico, Digital"),
        "publico_alvo": find_by_prefix("Publico alvo"),
        "b2b_b2c": find_by_prefix("B2B ou B2C"),
        "dores": find_by_prefix("Dores que resolve"),
        "modelo_produto": find_by_prefix("Modelo de Produto"),
        "escopo": find_by_prefix("Local ou Nacional"),
        "cidades": cidades,
        # Campo do briefing "entender o cliente": "Defina 2 palavras para SEO"
        "seo_keywords": find_by_prefix("Defina 2 palavras"),
        "termos_raiz": find_by_prefix("Termos raiz"),
    }
    # fallback do nicho generico (sem o prefixo "principal")
    if not briefing["nicho"]:
        # pega o primeiro campo cujo label seja exatamente "Nicho"
        for k, v in fields.items():
            if k.strip().lower() == "nicho":
                briefing["nicho"] = v
                break

    wb.close()
    return briefing


def keywords_from_briefing_seo(briefing: dict) -> list:
    """Extrai termos SEO do briefing (campo 'Defina 2 palavras para SEO').

    Aceita lista separada por virgula/ponto-e-virgula/pipe/quebra de linha.
    Retorna lista de dicts no mesmo formato de read_keywords_from_xlsx.
    """
    raw = (briefing.get("seo_keywords") or "").strip()
    if not raw:
        return []

    parts = [p.strip() for p in re.split(r"[,;|\n]+", raw) if p.strip()]
    seen = set()
    out = []
    for part in parts:
        key = part.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append({
            "keyword": part,
            "funil": "",
            "volume": 0,
            "position": None,
        })
    return out


def read_keywords_from_xlsx(xlsx_path: Path) -> list:
    """Le a aba 'Palavras-chave' e retorna lista de dicts ordenada por volume desc."""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    if "Palavras-chave" not in wb.sheetnames:
        wb.close()
        return []
    ws = wb["Palavras-chave"]
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(c or "").strip() for c in rows[0]]

    def col(name):
        for i, h in enumerate(headers):
            if h.lower().startswith(name.lower()):
                return i
        return None

    i_kw = col("Palavra")
    i_funil = col("Funil")
    i_vol = col("Volume")
    i_pos = col("Posicao")

    out = []
    for r in rows[1:]:
        if not r or i_kw is None or not r[i_kw]:
            continue
        vol = r[i_vol] if i_vol is not None else None
        out.append({
            "keyword": str(r[i_kw]).strip(),
            "funil": (str(r[i_funil]).strip() if (i_funil is not None and r[i_funil]) else ""),
            "volume": int(vol) if isinstance(vol, (int, float)) else 0,
            "position": r[i_pos] if i_pos is not None else None,
        })
    wb.close()
    out.sort(key=lambda x: x["volume"], reverse=True)
    return out


def read_similarity_from_xlsx(xlsx_path: Path) -> list:
    """Le a aba 'Similaridade Top 5' e retorna lista de dicts."""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    if "Similaridade Top 5" not in wb.sheetnames:
        wb.close()
        return []
    ws = wb["Similaridade Top 5"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        wb.close()
        return []
    headers = [str(c or "").strip() for c in rows[0]]

    def col(name):
        for i, h in enumerate(headers):
            if h.lower() == name.lower():
                return i
        return None

    i_kw = col("Keyword")
    i_vol = col("Volume")
    i_sim = col("Similaridade")

    out = []
    for r in rows[1:]:
        if not r or i_kw is None or not r[i_kw]:
            continue
        vol = r[i_vol] if i_vol is not None else None
        out.append({
            "keyword": str(r[i_kw]).strip(),
            "volume": int(vol) if isinstance(vol, (int, float)) else 0,
            "similaridade": str(r[i_sim]).strip() if (i_sim is not None and r[i_sim]) else "",
        })
    wb.close()
    return out


# ----------------------------------------------------------------------
# Selecao de keywords (3 + 2 distintas por intencao)
# ----------------------------------------------------------------------
def select_keywords(kws_list: list, sim_rows: list, brand: str = "") -> list:
    """Pega ate 3 da lista geral + ate 2 das que tem similaridade Alto,
    todas distintas por intent signature. Descarta keywords sem volume e de marca.
    """
    brand = (brand or "").strip().lower()

    def _is_brand(k):
        if not brand:
            return False
        return is_brand_keyword((k.get("keyword") or "").lower(), brand)

    # 1) top 3 com volume > 0 e nao-marca
    n_drop_brand = 0
    kws_with_vol = []
    for k in kws_list:
        if k.get("volume", 0) <= 0:
            continue
        if _is_brand(k):
            n_drop_brand += 1
            continue
        kws_with_vol.append(k)
    if n_drop_brand:
        print(f"[2]   {n_drop_brand} keyword(s) de marca descartada(s)")
    primary, dropped = dedupe_keywords_by_intent(kws_with_vol, MAX_KEYWORDS_KW)
    if dropped:
        print(f"[2]   {len(dropped)} variacao(oes) descartada(s) na escolha das primarias:")
        for d in dropped[:3]:
            print(f"        - '{d['keyword']}' (mesma intencao de '{d['duplicate_of']}')")

    # 2) keywords da similaridade com pelo menos 1 row "Alto" (descarta marca)
    alto_by_kw = {}
    for s in sim_rows:
        if s.get("similaridade") != "Alto":
            continue
        kw = s["keyword"]
        if brand and is_brand_keyword(kw.lower(), brand):
            continue
        if kw not in alto_by_kw:
            alto_by_kw[kw] = {"keyword": kw, "volume": s.get("volume") or 0}
    alto_kws = list(alto_by_kw.values())
    alto_kws.sort(key=lambda x: x["volume"], reverse=True)

    # 3) dedupe entre alto e primary
    primary_sigs = [keyword_signature(k["keyword"]) for k in primary]
    secondary = []
    sec_dropped = []
    for k in alto_kws:
        sig = keyword_signature(k["keyword"])
        if any(sig == s for s in primary_sigs):
            sec_dropped.append({"keyword": k["keyword"], "duplicate_of": "ja na lista primaria"})
            continue
        if any(sig == keyword_signature(s["keyword"]) for s in secondary):
            sec_dropped.append({"keyword": k["keyword"], "duplicate_of": "duplicada com outra Alto"})
            continue
        secondary.append(k)
        if len(secondary) >= MAX_KEYWORDS_SIM:
            break
    if sec_dropped:
        print(f"[2]   {len(sec_dropped)} keyword(s) Alto descartada(s):")
        for d in sec_dropped[:3]:
            print(f"        - '{d['keyword']}' ({d['duplicate_of']})")

    selected = primary + secondary
    selected = selected[:MAX_TOTAL_KEYWORDS]
    return selected


# ----------------------------------------------------------------------
# Consolida concorrentes unicos a partir das SERPs
# ----------------------------------------------------------------------
def consolidate_competitors(serps: dict, client_domain: str) -> list:
    """Junta todas as URLs do top 10 das varias SERPs, deduplica POR DOMINIO
    (mantem a melhor posicao + lista de keywords onde aparece) e remove o cliente
    e dominios ignorados."""
    by_domain = {}
    cli = normalize_domain(client_domain)
    for kw, items in serps.items():
        for it in items[:SERP_DEPTH_TOP10]:
            dom = normalize_domain(it.get("domain", ""))
            if not dom or dom == cli or dom in DOMAINS_TO_IGNORE:
                continue
            if dom not in by_domain:
                by_domain[dom] = {
                    "domain": dom,
                    "url": it["url"],
                    "title": it.get("title", ""),
                    "keywords": [],
                    "best_rank": it["rank"],
                }
            by_domain[dom]["keywords"].append({"kw": kw, "rank": it["rank"]})
            if it["rank"] < by_domain[dom]["best_rank"]:
                by_domain[dom]["best_rank"] = it["rank"]
                by_domain[dom]["url"] = it["url"]  # usa URL da melhor posicao
                by_domain[dom]["title"] = it.get("title", "")

    competitors = list(by_domain.values())
    competitors.sort(key=lambda x: (len(x["keywords"]), -x["best_rank"]), reverse=True)
    return competitors


# ----------------------------------------------------------------------
# Roda similaridade + perfil em paralelo
# ----------------------------------------------------------------------
def analyze_all_competitors(competitors: list, client_url: str, briefing: dict,
                             workers: int = 5) -> list:
    print(f"[4] Analisando similaridade/perfil de {len(competitors)} concorrentes (LLM, {workers} threads) ...")
    results = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        future_to_comp = {pool.submit(analyze_similarity, c["url"], client_url, briefing): c
                          for c in competitors}
        done = 0
        for fut in as_completed(future_to_comp):
            c = future_to_comp[fut]
            try:
                r = fut.result()
            except Exception as e:
                r = {"nicho": "—", "similaridade": "—", "perfil": "—", "motivo": f"Erro: {e}"}
            results.append({**c, **r})
            done += 1
            print(f"[4]   ({done}/{len(competitors)}) {r.get('similaridade','—')}/{r.get('perfil','—')}  {c['domain']}")
    # ordena: Alto primeiro, depois numero de aparicoes
    sim_rank = {"Alto": 3, "Medio": 2, "Baixo": 1, "—": 0}
    results.sort(key=lambda x: (sim_rank.get(x.get("similaridade", "—"), 0),
                                  len(x.get("keywords", []))), reverse=True)
    return results


# ----------------------------------------------------------------------
# Trafego Semrush em 3 periodos
# ----------------------------------------------------------------------
def pick_4_periods(series: list) -> dict:
    """Pega ponto mais recente + ~12m antes + ~24m antes + ~36m antes."""
    if not series:
        return {"current": None, "y1": None, "y2": None, "y3": None}
    current = series[-1]
    t1 = current["ts"] - 365 * 86400
    t2 = current["ts"] - 730 * 86400
    t3 = current["ts"] - 1095 * 86400

    def closest(target):
        return min(series, key=lambda p: abs(p["ts"] - target))

    y1 = closest(t1) if len(series) > 1 else None
    y2 = closest(t2) if len(series) > 12 else None
    y3 = closest(t3) if len(series) > 24 else None
    return {"current": current, "y1": y1, "y2": y2, "y3": y3}


def pick_3_periods(series: list) -> dict:
    """Compatibilidade retroativa para imports antigos."""
    return pick_4_periods(series)


def build_3period_traffic(client_url: str, competitors: list,
                          client_profile: str,
                          max_competitors: int = MAX_COMPETITORS_FOR_TRAFFIC) -> dict:
    cli = normalize_domain(client_url)
    # so concorrentes Alto vao para o relatorio de trafego (Medio/Baixo nao)
    alto = [c for c in competitors if c.get("similaridade") == "Alto"][:max_competitors]

    targets = [(cli, "cliente", client_profile)]
    for c in alto:
        targets.append((c["domain"], "concorrente", c.get("perfil", "—")))

    print(f"[5] Trafego Semrush 4 periodos: cliente + {len(alto)} concorrente(s) Alto")

    rows = []
    cur_dates, y1_dates, y2_dates, y3_dates = set(), set(), set(), set()
    for dom, role, perfil in targets:
        series = fetch_semrush_monthly_trend(dom)
        if not series:
            print(f"[5]   {dom}: sem dados")
            rows.append({"domain": dom, "role": role, "perfil": perfil,
                         "current": None, "y1": None, "y2": None, "y3": None,
                         "growth_seo_1y": None, "growth_seo_2y": None, "growth_seo_3y": None,
                         "growth_brand_1y": None, "growth_brand_2y": None, "growth_brand_3y": None})
            continue
        pts = pick_4_periods(series)
        cur, y1, y2, y3 = pts["current"], pts["y1"], pts["y2"], pts["y3"]

        def pct(now, then):
            if not then or then == 0:
                return None
            return round((now - then) / then * 100, 1)

        g1 = pct(cur["non_branded"], y1["non_branded"]) if (cur and y1) else None
        g2 = pct(cur["non_branded"], y2["non_branded"]) if (cur and y2) else None
        g3 = pct(cur["non_branded"], y3["non_branded"]) if (cur and y3) else None
        b1 = pct(cur["branded"], y1["branded"]) if (cur and y1) else None
        b2 = pct(cur["branded"], y2["branded"]) if (cur and y2) else None
        b3 = pct(cur["branded"], y3["branded"]) if (cur and y3) else None

        print(f"[5]   {dom}: "
              f"SEO {cur['date']}={cur['non_branded']:,} "
              f"vs {y1['date'] if y1 else '?'}={y1['non_branded'] if y1 else '?':,}"
              f"  (cresc 1y={g1}%)" if y1 else "")

        if cur: cur_dates.add(cur["date"])
        if y1: y1_dates.add(y1["date"])
        if y2: y2_dates.add(y2["date"])
        if y3: y3_dates.add(y3["date"])
        rows.append({
            "domain": dom, "role": role, "perfil": perfil,
            "current": cur, "y1": y1, "y2": y2, "y3": y3,
            "growth_seo_1y": g1, "growth_seo_2y": g2, "growth_seo_3y": g3,
            "growth_brand_1y": b1, "growth_brand_2y": b2, "growth_brand_3y": b3,
        })

    return {
        "rows": rows,
        "current_date": ", ".join(sorted(cur_dates)) if cur_dates else "",
        "y1_date": ", ".join(sorted(y1_dates)) if y1_dates else "",
        "y2_date": ", ".join(sorted(y2_dates)) if y2_dates else "",
        "y3_date": ", ".join(sorted(y3_dates)) if y3_dates else "",
    }


# ----------------------------------------------------------------------
# XLSX final
# ----------------------------------------------------------------------
HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
THIN = Side(style="thin", color="CBD5E1")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FILL_SIM = {
    "Alto":  PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid"),
    "Medio": PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
    "Baixo": PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"),
}
FILL_PERFIL = {
    "Especialista": PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"),
    "Generalista":  PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
}


def write_xlsx(out_path: Path, client_input: str, briefing: dict,
               selected_kws: list, competitors_full: list, traffic: dict):
    wb = Workbook()

    # ---- Aba "Resumo"
    ws = wb.active
    ws.title = "Resumo"
    ws.append(["Concorrentes Finder"])
    ws.merge_cells("A1:B1")
    ws["A1"].font = Font(bold=True, size=14, color="0F172A")
    ws.append([])
    ws.append(["Cliente (input)", client_input])
    ws.append(["URL cliente", briefing.get("url", "")])
    ws.append(["Nicho principal", briefing.get("nicho_principal", "")])
    ws.append(["Escopo", briefing.get("escopo", "")])
    ws.append(["Data", datetime.now().strftime("%d/%m/%Y %H:%M")])
    ws.append([])
    ws.append(["Keywords usadas para varrer SERP"])
    ws.append(["#", "Keyword", "Volume", "Origem"])
    for c in ws[9]:
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
    for i, k in enumerate(selected_kws, 1):
        ws.append([i, k["keyword"], k.get("volume"), k.get("source", "")])

    ws.append([])
    ws.append(["Concorrentes encontrados"])
    ws.append(["Total unico", len(competitors_full)])
    sim_counts = {"Alto": 0, "Medio": 0, "Baixo": 0, "—": 0}
    for c in competitors_full:
        sim_counts[c.get("similaridade", "—")] = sim_counts.get(c.get("similaridade", "—"), 0) + 1
    ws.append(["Similaridade Alto", sim_counts.get("Alto", 0)])
    ws.append(["Similaridade Medio", sim_counts.get("Medio", 0)])
    ws.append(["Similaridade Baixo", sim_counts.get("Baixo", 0)])

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 18

    # ---- Aba "Concorrentes"
    ws_c = wb.create_sheet("Concorrentes")
    headers = ["#", "Dominio", "Similaridade", "Perfil", "Nicho", "Apareceu em (kws)",
               "Melhor Posicao", "URL", "Titulo", "Motivo"]
    ws_c.append(headers)
    for c in ws_c[1]:
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    for i, comp in enumerate(competitors_full, 1):
        kws_str = ", ".join(f"{k['kw']} (#{k['rank']})" for k in comp.get("keywords", []))
        ws_c.append([
            i, comp["domain"], comp.get("similaridade", ""), comp.get("perfil", ""),
            comp.get("nicho", ""), kws_str, comp.get("best_rank", ""),
            comp.get("url", ""), comp.get("title", ""), comp.get("motivo", ""),
        ])
        last = ws_c.max_row
        f = FILL_SIM.get(comp.get("similaridade", ""))
        if f:
            ws_c.cell(row=last, column=3).fill = f
        fp = FILL_PERFIL.get(comp.get("perfil", ""))
        if fp:
            ws_c.cell(row=last, column=4).fill = fp
        for cell in ws_c[last]:
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    widths = [4, 28, 14, 14, 28, 50, 12, 50, 38, 60]
    for i, w in enumerate(widths, 1):
        ws_c.column_dimensions[chr(64 + i)].width = w

    # ---- Aba "Trafego SEO 3 periodos" (cliente + Alto)
    if traffic and traffic.get("rows"):
        ws_t = wb.create_sheet("Trafego SEO Marca")
        cur = traffic.get("current_date") or "Atual"
        y1 = traffic.get("y1_date") or "12m antes"
        y2 = traffic.get("y2_date") or "24m antes"
        ws_t.append(["Trafego organico (Semrush): SEO non-branded + Marca - 3 periodos"])
        ws_t.merge_cells("A1:M1")
        ws_t["A1"].font = Font(bold=True, size=12, color="0F172A")
        ws_t.append([])
        # 2 niveis de cabecalho
        ws_t.append([
            "Tipo", "Dominio", "Perfil",
            f"SEO {cur}", f"SEO {y1}", f"SEO {y2}",
            f"Marca {cur}", f"Marca {y1}", f"Marca {y2}",
            f"Total {cur}", f"Total {y1}", f"Total {y2}",
            "Cresc SEO 1a (%)", "Cresc SEO 2a (%)",
        ])
        # ajusta merge da row 1
        ws_t.unmerge_cells("A1:M1")
        ws_t.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
        ws_t["A1"].font = Font(bold=True, size=12, color="0F172A")
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
def main():
    if len(sys.argv) < 2:
        print("Uso: python find_concorrentes.py <nome_cliente>")
        print("Ex:  python find_concorrentes.py cnempreendimentos")
        print("     python find_concorrentes.py apetsaude")
        sys.exit(1)

    client_input = sys.argv[1].strip()
    print(f"\n=== Concorrentes Finder: '{client_input}' ===\n")
    t0 = time.time()

    # 1) localiza e le XLSX do cliente
    xlsx_path = find_client_xlsx(client_input)
    briefing = read_briefing_from_xlsx(xlsx_path)
    kws_list = read_keywords_from_xlsx(xlsx_path)
    sim_rows = read_similarity_from_xlsx(xlsx_path)
    print(f"[1] Briefing: nicho='{briefing['nicho_principal']}', escopo={briefing['escopo']}, "
          f"url={briefing.get('url','')}")
    print(f"[1] Lidas {len(kws_list)} keywords e {len(sim_rows)} linhas de similaridade")

    if not briefing.get("url"):
        print("ERRO: URL do cliente nao encontrada no briefing.")
        sys.exit(1)

    # 2) seleciona 3 + 2 keywords (sem duplicatas por intent, sem marca, com volume > 0)
    client_brand = brand_token_from_domain(briefing.get("url", ""))
    selected = select_keywords(kws_list, sim_rows, brand=client_brand)
    # marca origem (primary = aba palavras-chave; alto = aba similaridade)
    for i, k in enumerate(selected):
        k["source"] = "palavras-chave (top vol)" if i < MAX_KEYWORDS_KW else "similaridade Alto"
    print(f"[2] {len(selected)} keyword(s) selecionada(s):")
    for k in selected:
        print(f"      - '{k['keyword']}' (vol {k['volume']}, {k['source']})")

    if not selected:
        print("ERRO: nenhuma keyword selecionada.")
        sys.exit(1)

    # 3) top 10 SERPs via DataForSEO
    keywords_only = [k["keyword"] for k in selected]
    print(f"[3] Buscando SERPs top {SERP_DEPTH_TOP10} no DataForSEO ({len(keywords_only)} keywords) ...")
    serps = fetch_serps(keywords_only, depth=SERP_DEPTH_TOP10, workers=5)

    # 4) consolida concorrentes
    competitors = consolidate_competitors(serps, briefing["url"])
    print(f"[3] {len(competitors)} concorrentes unicos no top {SERP_DEPTH_TOP10}")

    # 5) similaridade + perfil
    competitors_full = analyze_all_competitors(competitors, briefing["url"], briefing,
                                                 workers=5)

    # 6) perfil do cliente
    print("[5] Classificando perfil do cliente ...")
    client_url_full = briefing["url"] if briefing["url"].startswith(("http://", "https://")) \
        else f"https://{briefing['url']}"
    try:
        site = scrape_site(client_url_full)
        client_profile = classify_client_profile(client_url_full, site, briefing)
    except Exception as e:
        print(f"[5] Falha ao classificar cliente: {e}")
        client_profile = "—"
    print(f"[5] Cliente classificado como: {client_profile}")

    # 7) trafego Semrush 3 periodos
    traffic = build_3period_traffic(client_url_full, competitors_full, client_profile)

    # 8) saida
    client_slug = xlsx_path.parent.name
    out_path = OUT_DIR / client_slug / f"concorrentes-{client_slug}.xlsx"
    print(f"[OUT] Escrevendo: {out_path}")
    saved = write_xlsx(out_path, client_input, briefing, selected, competitors_full, traffic)
    print(f"[OUT] XLSX salvo: {saved}")

    print(f"\n=== Concluido em {time.time()-t0:.1f}s ===")
    print(f"XLSX: {saved}")


if __name__ == "__main__":
    main()
