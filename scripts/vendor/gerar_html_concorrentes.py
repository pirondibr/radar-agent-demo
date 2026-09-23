# -*- coding: utf-8 -*-
"""
Gera HTML com abas a partir do XLSX unificado de concorrentes.

Uso:
    python gerar_html_concorrentes.py grupodata
    python gerar_html_concorrentes.py caminho/concorrentes-all-grupodata.xlsx
"""

from __future__ import annotations

import html
import re
import sys
from pathlib import Path

import openpyxl

_THIS_DIR = Path(__file__).resolve().parent
OUT_DIR = _THIS_DIR / "output"

TAB_ORDER = [
    "Resumo",
    "Resumo Execucao",
    "Resumo Anunciantes",
    "Filtro - Concorrentes",
    "Anuncios",
    "Amostra por pagante",
    "Concorrentes SEO",
    "Concorrentes Maps",
    "Concorrentes Ads",
    "Concorrentes LLM",
    "Concorrentes Unificado",
    "Metricas Canais",
    "Trafego SEO Marca",
    "Crescimento",
]
TAB_LABELS = {
    "Resumo": "Resumo",
    "Resumo Execucao": "Resumo",
    "Resumo Anunciantes": "Resumo Anunciantes",
    "Filtro - Concorrentes": "Filtro Concorrentes",
    "Anuncios": "Anuncios",
    "Amostra por pagante": "Amostra por pagante",
    "Concorrentes SEO": "SEO",
    "Concorrentes Maps": "Maps",
    "Concorrentes Ads": "Ads",
    "Concorrentes LLM": "LLM",
    "Concorrentes Unificado": "Unificado",
    "Metricas Canais": "Metricas Canais",
    "Trafego SEO Marca": "Trafego",
    "Crescimento": "Crescimento",
}


def _esc(value) -> str:
    if value is None:
        return ""
    return html.escape(str(value), quote=True)


def _cell(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _fold(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def find_xlsx(client_input: str) -> Path:
    raw = client_input.strip()
    p = Path(raw)
    if p.exists() and p.suffix.lower() == ".xlsx":
        return p
    slug = re.sub(r"[^a-z0-9]+", "", raw.lower().replace("https://", "").replace("www.", "").split("/")[0].split(".")[0])
    folder = OUT_DIR / slug
    files = sorted(folder.glob(f"concorrentes-all-{slug}*.xlsx"),
                   key=lambda x: x.stat().st_mtime, reverse=True)
    if not files:
        files = sorted(folder.glob("concorrentes-all-*.xlsx"),
                       key=lambda x: x.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError(
            f"XLSX de concorrentes nao encontrado em {folder}.\n"
            f"Rode primeiro: python find_concorrentes_all.py {slug or raw}"
        )
    return files[0]


def load_sheets(xlsx_path: Path) -> dict[str, list[list]]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    out = {}
    for name in wb.sheetnames:
        ws = wb[name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append([_cell(c) for c in row])
        out[name] = rows
    wb.close()
    return out


def parse_resumo(rows: list[list]) -> dict:
    meta = {}
    sections = []
    current = None
    for row in rows:
        a = (row[0] if row else "").strip()
        b = (row[1] if len(row) > 1 else "").strip()
        if not a:
            continue
        is_header = a.upper().startswith("MODO") or a.upper() in ("UNIFICADO",) or a.startswith("Concorrentes Finder")
        if is_header and a.startswith("Concorrentes Finder"):
            meta["titulo"] = a
            continue
        if is_header:
            current = {"title": a, "items": []}
            sections.append(current)
            continue
        if current is not None:
            current["items"].append((a, b))
        else:
            meta[a] = b
    return {"meta": meta, "sections": sections}


def _lv_class(val: str) -> str:
    s = (val or "").strip()
    if not s:
        return ""
    key = s.replace("é", "e").replace("á", "a")
    return f"lv-{key}"


def _looks_url(val: str) -> bool:
    v = (val or "").lower()
    return v.startswith("http://") or v.startswith("https://")


def _todos_hidden_header(hl: str) -> bool:
    return hl in {
        "similaridade", "perfil", "fonte",
        "meta ads url", "google ads url", "linkedin ads url",
        "instagram url", "youtube url", "tiktok url",
        "url", "url trafego", "origem url", "keywords", "motivo",
    }


def _format_num(val: str) -> str:
    s = (val or "").strip()
    if not s:
        return "—"
    try:
        n = float(s.replace(",", ""))
        if abs(n - int(n)) < 1e-9:
            return f"{int(n):,}".replace(",", ".")
        return f"{n:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except ValueError:
        return s


def _format_pct(val: str) -> str:
    s = (val or "").strip()
    if not s:
        return "—"
    try:
        n = float(s)
        cls = "pos" if n >= 0 else "neg"
        sign = "+" if n > 0 else ""
        return f'<span class="pct {cls}">{sign}{n:.1f}%</span>'
    except ValueError:
        return _esc(s)


def _sort_attr(val: str) -> str:
    raw_num = _raw_num(val)
    return raw_num if raw_num else _fold(val)


def render_competitor_table(sheet_name: str, rows: list[list], tab_id: str) -> str:
    if len(rows) < 2:
        return '<p class="empty">Nenhum concorrente nesta aba.</p>'
    raw_headers = [h or f"Col{i}" for i, h in enumerate(rows[0], 1)]
    raw_headers_l = [_fold(h) for h in raw_headers]
    display_idx = list(range(len(raw_headers)))
    if sheet_name == "Metricas Canais":
        display_idx = [i for i, h in enumerate(raw_headers_l) if not (h == "url" or h.endswith(" url"))]
    headers = [raw_headers[i] for i in display_idx]
    headers_l = [raw_headers_l[i] for i in display_idx]
    hidden_flags = [_todos_hidden_header(h) if sheet_name == "Todos" else False for h in headers_l]
    i_sim = next((i for i, h in enumerate(headers_l) if h.startswith("similaridade")), None)
    i_perfil = next((i for i, h in enumerate(headers_l) if h == "perfil"), None)
    i_url = next((i for i, h in enumerate(raw_headers_l) if h == "url"), None)
    i_fonte = next((i for i, h in enumerate(headers_l) if h == "fonte"), None)

    sims = sorted({r[i_sim] for r in rows[1:] if i_sim is not None and i_sim < len(r) and r[i_sim]})
    fontes = sorted({r[i_fonte] for r in rows[1:] if i_fonte is not None and i_fonte < len(r) and r[i_fonte]})

    filters = [f'<div class="filters" data-table="{_esc(tab_id)}">']
    if sims:
        opts = "".join(f'<option value="{_esc(s)}">{_esc(s)}</option>' for s in sims)
        filters.append(
            f'<label>Similaridade <select class="f-sim"><option value="">Todas</option>{opts}</select></label>'
        )
    if fontes:
        opts = "".join(f'<option value="{_esc(s)}">{_esc(s)}</option>' for s in fontes)
        filters.append(
            f'<label>Fonte <select class="f-fonte"><option value="">Todas</option>{opts}</select></label>'
        )
    filters.append(
        f'<label>Busca <input type="search" class="f-q" placeholder="dominio, nicho, motivo..."></label>'
    )
    if sheet_name == "Todos" and any(hidden_flags):
        filters.append('<button type="button" class="toggle-extra-cols" data-show="0">Mostrar colunas extras</button>')
    filters.append("</div>")

    thead = "".join(
        f'<th class="sortable{" extra-col is-hidden-col" if hidden_flags[i] else ""}" data-col="{i}">{_esc(h)}<span class="sort-ind"></span></th>'
        for i, h in enumerate(headers)
    )
    body = []
    for r in rows[1:]:
        if not any(r):
            continue
        row_display = [r[i] if i < len(r) else "" for i in display_idx]
        sim = row_display[i_sim] if i_sim is not None and i_sim < len(row_display) else ""
        fonte = row_display[i_fonte] if i_fonte is not None and i_fonte < len(row_display) else ""
        tds = []
        for i, h in enumerate(headers):
            val = row_display[i] if i < len(row_display) else ""
            hl = headers_l[i]
            sort_attr = _sort_attr(val)
            extra_cls = " extra-col is-hidden-col" if hidden_flags[i] else ""
            def td_wrap(inner: str, cls: str = "") -> str:
                full_cls = (cls + extra_cls).strip()
                cls_attr = f' class="{full_cls}"' if full_cls else ""
                return f'<td{cls_attr} data-n="{_esc(sort_attr)}">{inner}</td>'
            def row_val(header_name: str) -> str:
                try:
                    j = raw_headers_l.index(header_name)
                except ValueError:
                    return ""
                return r[j] if j < len(r) else ""
            if hl.startswith("similaridade") or hl == "perfil":
                tds.append(td_wrap(f'<span class="{_lv_class(val)}">{_esc(val) or "—"}</span>', "c"))
            elif hl == "quantidade de anuncios":
                href = row_val("linkedin ad library") or row_val("meta ad library")
                tds.append(td_wrap(f'<a href="{_esc(href)}" target="_blank" rel="noopener">{_esc(val) or "—"}</a>' if href and _looks_url(href) else (_esc(val) or "—")))
            elif hl == "meta ads":
                href = row_val("meta ads url")
                tds.append(td_wrap(f'<a href="{_esc(href)}" target="_blank" rel="noopener">{_esc(val) or "—"}</a>' if href and _looks_url(href) else (_esc(val) or "—")))
            elif hl == "google ads":
                href = row_val("google ads url")
                tds.append(td_wrap(f'<a href="{_esc(href)}" target="_blank" rel="noopener">{_esc(val) or "—"}</a>' if href and _looks_url(href) else (_esc(val) or "—")))
            elif hl == "linkedin ads":
                href = row_val("linkedin ads url")
                tds.append(td_wrap(f'<a href="{_esc(href)}" target="_blank" rel="noopener">{_esc(val) or "—"}</a>' if href and _looks_url(href) else (_esc(val) or "—")))
            elif hl == "instagram seguidores":
                href = row_val("instagram url")
                tds.append(td_wrap(f'<a href="{_esc(href)}" target="_blank" rel="noopener">{_esc(val) or "—"}</a>' if href and _looks_url(href) else (_esc(val) or "—")))
            elif hl == "youtube seguidores":
                href = row_val("youtube url")
                tds.append(td_wrap(f'<a href="{_esc(href)}" target="_blank" rel="noopener">{_esc(val) or "—"}</a>' if href and _looks_url(href) else (_esc(val) or "—")))
            elif hl == "tiktok seguidores":
                href = row_val("tiktok url")
                tds.append(td_wrap(f'<a href="{_esc(href)}" target="_blank" rel="noopener">{_esc(val) or "—"}</a>' if href and _looks_url(href) else (_esc(val) or "—")))
            elif (hl == "url" or hl.endswith(" url")) and _looks_url(val):
                tds.append(td_wrap(f'<a href="{_esc(val)}" target="_blank" rel="noopener">{_esc(val)}</a>'))
            elif hl.startswith("dominio"):
                url = r[i_url] if i_url is not None and i_url < len(r) and _looks_url(r[i_url]) else ""
                if not url and val:
                    url = "https://" + val
                label = _esc(val)
                if len(val or "") > 26:
                    label = f'<span title="{_esc(val)}">{_esc(val[:26])}...</span>'
                tds.append(td_wrap(f'<a class="dom" href="{_esc(url)}" target="_blank" rel="noopener">{label}</a>', "clip"))
            elif hl == "fonte":
                tags = "".join(f'<span class="fonte">{_esc(p.strip())}</span>' for p in val.split("+") if p.strip())
                tds.append(td_wrap(tags or _esc(val)))
            elif hl == "motivo":
                full = val or ""
                cut = max(80, len(full) // 2) if full else 0
                short = full[:cut].rstrip()
                if short and len(short) < len(full):
                    short += "..."
                tds.append(td_wrap(f'<span title="{_esc(full)}">{_esc(short) or "—"}</span>', "clip"))
            else:
                tds.append(td_wrap(_esc(val)))
        body.append(
            f'<tr data-sim="{_esc(sim)}" data-fonte="{_esc(fonte)}">' + "".join(tds) + "</tr>"
        )

    n = len(body)
    return (
        f'<div class="tab-head"><span class="count">{n} linha(s)</span></div>'
        + "".join(filters)
        + f'<div class="table-wrap"><table class="data-table sortable-generic" id="tbl-{_esc(tab_id)}">'
        f"<thead><tr>{thead}</tr></thead><tbody>{''.join(body)}</tbody></table></div>"
    )


def _raw_num(val: str) -> str:
    s = (val or "").strip()
    if not s:
        return ""
    try:
        return str(float(s.replace(",", "")))
    except ValueError:
        return ""


def _age_sort_val(val: str) -> str:
    s = (val or "").strip()
    if not s:
        return ""
    if s.endswith("+"):
        s = s[:-1]
    return _raw_num(s)


def render_traffic(rows: list[list], hide_tipo: bool = False, swap_growth: bool = False) -> str:
    if not rows:
        return '<p class="empty">Sem dados de trafego Semrush.</p>'
    title = rows[0][0] if rows and rows[0] else "Trafego SEO"
    header_idx = None
    for i, r in enumerate(rows):
        if r and _fold(r[0]) in ("tipo",):
            header_idx = i
            break
    if header_idx is None:
        return f'<p class="empty">{_esc(title)}</p>'
    headers = rows[header_idx]
    last_header_idx = max((i for i, h in enumerate(headers) if str(h or "").strip()), default=-1)
    headers = headers[:last_header_idx + 1]
    body_rows = [r[:last_header_idx + 1] for r in rows[header_idx + 1:]]
    keep_idx = list(range(len(headers)))
    if hide_tipo and headers and _fold(headers[0]) == "tipo":
        keep_idx = keep_idx[1:]
        headers = [headers[i] for i in keep_idx]
        body_rows = [[r[i] if i < len(r) else "" for i in keep_idx] for r in body_rows]
    thead = "".join(
        f'<th class="sortable" data-col="{i}" title="Clique para ordenar maior/menor">{_esc(h)}'
        f'<span class="sort-ind"></span></th>'
        for i, h in enumerate(headers)
    )
    body = []
    for r in body_rows:
        if not any(r):
            continue
        role = (r[0] if r else "").lower()
        cls = ' class="cliente"' if role.startswith("cliente") else ""
        tds = []
        for i, val in enumerate(r):
            h = _fold(headers[i] if i < len(headers) else "")
            raw = _raw_num(val)
            if "cresc" in h or "%" in (headers[i] if i < len(headers) else ""):
                tds.append(f'<td class="c" data-n="{_esc(raw)}">{_format_pct(val)}</td>')
            elif h == "idade":
                tds.append(f'<td class="c" data-n="{_esc(_age_sort_val(val))}">{_esc(val) or "—"}</td>')
            elif h in ("tipo", "dominio", "perfil"):
                if h == "perfil":
                    tds.append(f'<td class="c" data-n="{_esc(val)}"><span class="{_lv_class(val)}">{_esc(val) or "—"}</span></td>')
                elif h == "dominio":
                    tds.append(f'<td data-n="{_esc(val)}"><span class="dom">{_esc(val)}</span></td>')
                else:
                    tds.append(f'<td data-n="{_esc(val)}">{_esc(val)}</td>')
            else:
                tds.append(f'<td class="num" data-n="{_esc(raw)}">{_format_num(val)}</td>')
        body.append(f"<tr{cls}>" + "".join(tds) + "</tr>")
    controls = ""
    table_id = "tbl-trafego"
    table_class = "data-table traffic"
    if swap_growth:
        controls = (
            '<div class="swap-toolbar">'
            '<button type="button" class="swap-btn" data-growth-toggle="seo">SEO primeiro</button>'
            '<button type="button" class="swap-btn ghost" data-growth-toggle="marca">Marca primeiro</button>'
            "</div>"
        )
        table_id = "tbl-crescimento"
        table_class += " growth-swap"
    return (
        f'<p class="note">{_esc(title)} &middot; Clique no cabecalho da coluna para ordenar do maior para o menor.</p>'
        + controls
        + f'<div class="table-wrap"><table class="{table_class}" id="{table_id}">'
        f"<thead><tr>{thead}</tr></thead><tbody>{''.join(body)}</tbody></table></div>"
    )


def render_resumo(parsed: dict, sheets: dict) -> str:
    meta = parsed["meta"]
    url = meta.get("URL cliente") or meta.get("URL") or ""
    nicho = meta.get("Nicho principal") or meta.get("Nicho") or "—"
    escopo = meta.get("Escopo", "—")
    perfil = meta.get("Perfil do cliente") or meta.get("Perfil") or "—"
    data = meta.get("Data", "")
    cliente = meta.get("Cliente (input)") or meta.get("Cliente") or ""

    counts = {}
    for name, rows in sheets.items():
        if name.startswith("Concorrentes"):
            counts[name] = max(0, len([r for r in rows[1:] if any(r)]))
    unified_count = counts.get("Concorrentes Unificado")
    if unified_count is None and "Resumo Anunciantes" in sheets:
        unified_count = max(0, len([r for r in sheets["Resumo Anunciantes"][1:] if any(r)]))
    if unified_count is None:
        unified_count = 0

    cards = [
        ("Nicho", nicho, "c1"),
        ("Escopo", escopo, "c2"),
        ("Perfil", perfil, "c3"),
        ("Unificado", str(unified_count), "c4"),
    ]
    cards_html = "".join(
        f'<div class="card {cls}"><div class="label">{_esc(lab)}</div><div class="val">{_esc(val)}</div></div>'
        for lab, val, cls in cards
    )

    link = f'<a href="{_esc(url)}" target="_blank" rel="noopener">{_esc(url)}</a>' if url else "—"
    meta_html = (
        '<div class="info-row">'
        f"<span>Cliente: <b>{_esc(cliente)}</b></span>"
        f"<span>URL: <b>{link}</b></span>"
        f"<span>Cidade(s): <b>{_esc(meta.get('Cidade(s)') or meta.get('Cidades') or '—')}</b></span>"
        f"<span>Data: <b>{_esc(data)}</b></span>"
        "</div>"
    )

    sections = []
    for sec in parsed["sections"]:
        items = []
        for a, b in sec["items"]:
            items.append(f'<div class="kv"><div class="k">{_esc(a)}</div><div class="v">{_esc(b) or "—"}</div></div>')
        sections.append(
            f'<div class="block"><h3>{_esc(sec["title"])}</h3><div class="kv-grid">{"".join(items)}</div></div>'
        )

    maps_addr = ""
    maps_title = ""
    for sec in parsed["sections"]:
        if "Maps" in sec["title"]:
            for a, b in sec["items"]:
                if a.startswith("Maps - Endereco"):
                    maps_addr = b
                if a.startswith("Maps - Empresa"):
                    maps_title = b
    warning = ""
    addr_l = maps_addr.lower()
    if maps_addr and any(x in addr_l for x in ("espanha", "españa", "spain", "portugal", "usa", "united states")):
        warning = (
            f'<div class="warn"><b>Alerta Maps:</b> o match parece ser de outra empresa/pais. '
            f'Empresa: {_esc(maps_title)} · Endereco: {_esc(maps_addr)}. Ignore a aba Maps neste caso.</div>'
        )

    return (
        f'<div class="banner"><div class="lab">Nicho principal</div>'
        f'<div class="banner-val">{_esc(nicho)}</div></div>'
        + meta_html
        + f'<div class="cards">{cards_html}</div>'
        + warning
        + "".join(sections)
    )


def build_html(xlsx_path: Path, sheets: dict) -> str:
    slug = xlsx_path.parent.name
    parsed = parse_resumo(sheets.get("Resumo") or sheets.get("Resumo Execucao") or [])
    meta = parsed["meta"]
    url = meta.get("URL cliente", "")
    nicho = meta.get("Nicho principal", slug)

    tabs = []
    panels = []
    first = True
    ordered_sheet_names = [name for name in TAB_ORDER if name in sheets]
    for name in sheets:
        if name not in ordered_sheet_names:
            ordered_sheet_names.append(name)
    for name in ordered_sheet_names:
        tid = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")
        label = TAB_LABELS.get(name, name)
        n = ""
        if name.startswith("Concorrentes") or name in ("Resumo Anunciantes", "Filtro - Concorrentes", "Anuncios", "Amostra por pagante"):
            n_rows = len([r for r in sheets[name][1:] if any(r)])
            n = f" <em>{n_rows}</em>"
        tabs.append(
            f'<button type="button" class="tab{" active" if first else ""}" data-tab="{tid}">{_esc(label)}{n}</button>'
        )
        if name in ("Resumo", "Resumo Execucao"):
            inner = render_resumo(parsed, sheets)
        elif name.startswith("Trafego"):
            inner = render_traffic(sheets[name])
        elif name == "Crescimento":
            inner = render_traffic(sheets[name], hide_tipo=True, swap_growth=True)
        else:
            inner = render_competitor_table(name, sheets[name], tid)
        panels.append(
            f'<section class="panel{" active" if first else ""}" id="panel-{tid}">{inner}</section>'
        )
        first = False

    return f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Concorrentes — {_esc(slug)}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,sans-serif;background:#f8fafc;color:#1e293b;line-height:1.5}}
.wrap{{max-width:1280px;margin:0 auto;padding:32px 20px 48px}}
.brand-bar{{height:5px;background:linear-gradient(90deg,#0f766e,#2563eb,#7c3aed);border-radius:4px;margin-bottom:24px}}
h1{{font-size:26px;font-weight:800;color:#0f172a;margin-bottom:4px}}
.sub{{font-size:14px;color:#64748b;margin-bottom:16px}}
.info-row{{display:flex;flex-wrap:wrap;gap:8px 20px;font-size:13px;color:#475569;margin-bottom:18px}}
.info-row b{{color:#0f172a}}
.info-row a{{color:#2563eb;text-decoration:none}}
.banner{{background:linear-gradient(135deg,#0f766e,#2563eb);color:#fff;border-radius:12px;padding:20px 24px;margin-bottom:20px}}
.banner .lab{{font-size:11px;font-weight:700;letter-spacing:.8px;text-transform:uppercase;opacity:.85;margin-bottom:6px}}
.banner-val{{font-size:26px;font-weight:800}}
.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:22px}}
.card{{background:#fff;border-radius:12px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,.06);border-left:4px solid #2563eb}}
.card.c2{{border-left-color:#0f766e}}.card.c3{{border-left-color:#7c3aed}}.card.c4{{border-left-color:#f97316}}
.card .label{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#64748b;margin-bottom:4px}}
.card .val{{font-size:20px;font-weight:800;color:#0f172a}}
.tabs{{display:flex;flex-wrap:wrap;gap:6px;margin:8px 0 18px;border-bottom:2px solid #e2e8f0;padding-bottom:8px}}
.tab{{border:0;background:#f1f5f9;color:#334155;padding:8px 14px;border-radius:8px 8px 0 0;font-weight:700;font-size:13px;cursor:pointer}}
.tab em{{font-style:normal;background:#e2e8f0;border-radius:999px;padding:1px 7px;margin-left:6px;font-size:11px}}
.tab.active{{background:#2563eb;color:#fff}}
.tab.active em{{background:rgba(255,255,255,.22);color:#fff}}
.panel{{display:none}}
.panel.active{{display:block}}
.block{{background:#fff;border-radius:12px;padding:16px 18px;margin-bottom:14px;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.block h3{{font-size:14px;font-weight:800;color:#0f766e;margin-bottom:10px}}
.kv-grid{{display:grid;grid-template-columns:220px 1fr;gap:4px 16px}}
.kv .k{{font-size:12px;color:#64748b}}
.kv .v{{font-size:13px;color:#0f172a}}
.warn{{background:#fef3c7;border:1px solid #f59e0b;color:#92400e;padding:12px 14px;border-radius:10px;margin-bottom:16px;font-size:13px}}
.filters{{display:flex;flex-wrap:wrap;gap:10px;margin:0 0 12px;background:#fff;padding:10px 12px;border-radius:10px}}
.filters label{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#64748b;display:flex;flex-direction:column;gap:4px}}
.filters select,.filters input{{font-size:13px;padding:6px 8px;border:1px solid #e2e8f0;border-radius:8px;min-width:150px}}
.toggle-extra-cols{{border:0;background:#e2e8f0;color:#334155;padding:8px 12px;border-radius:8px;font-size:12px;font-weight:700;cursor:pointer}}
.tab-head{{margin-bottom:8px;font-size:12px;color:#64748b}}
.table-wrap{{overflow:auto;background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.data-table{{width:100%;border-collapse:collapse;font-size:12.5px}}
.data-table th{{background:#f1f5f9;color:#475569;font-size:10px;text-transform:uppercase;letter-spacing:.4px;text-align:left;padding:10px 10px;white-space:nowrap}}
.data-table td{{padding:9px 10px;border-top:1px solid #f1f5f9;vertical-align:top;color:#334155}}
.data-table td.c{{text-align:center}}
.data-table td.num{{text-align:right;font-variant-numeric:tabular-nums}}
.data-table td.clip{{max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.data-table .is-hidden-col{{display:none}}
.data-table tbody tr:hover{{background:#f8fafc}}
.data-table tbody tr.cliente{{background:#dbeafe;font-weight:700}}
.dom{{font-weight:700;color:#0f172a;text-decoration:none}}
.dom:hover{{color:#2563eb}}
.lv-Alto,.lv-Especialista{{display:inline-block;background:rgba(16,185,129,.14);color:#047857;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700;min-width:70px;text-align:center}}
.lv-Medio,.lv-Generalista{{display:inline-block;background:rgba(245,158,11,.14);color:#a16207;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700;min-width:70px;text-align:center}}
.lv-Baixo{{display:inline-block;background:rgba(220,38,38,.10);color:#b91c1c;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700;min-width:70px;text-align:center}}
.lv-Nenhum{{display:inline-block;background:rgba(100,116,139,.14);color:#475569;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700;min-width:70px;text-align:center}}
.fonte{{display:inline-block;background:#eef2ff;color:#3730a3;padding:2px 8px;border-radius:6px;font-size:11px;font-weight:700;margin:1px}}
.data-table th.sortable{{cursor:pointer;user-select:none;white-space:nowrap}}
.data-table th.sortable:hover{{color:#0f172a;background:#e2e8f0}}
.data-table th.sortable .sort-ind{{display:inline-block;min-width:12px;margin-left:4px;color:#94a3b8;font-size:10px}}
.data-table th.sorted-desc .sort-ind{{color:#2563eb}}
.data-table th.sorted-desc .sort-ind::after{{content:"▼"}}
.data-table th.sorted-asc .sort-ind{{color:#2563eb}}
.data-table th.sorted-asc .sort-ind::after{{content:"▲"}}
.pct.pos{{color:#16a34a;font-weight:700}}
.pct.neg{{color:#dc2626;font-weight:700}}
.note{{font-size:13px;color:#64748b;margin-bottom:10px}}
.swap-toolbar{{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 10px}}
.swap-btn{{border:0;background:#2563eb;color:#fff;padding:8px 12px;border-radius:8px;font-size:12px;font-weight:700;cursor:pointer}}
.swap-btn.ghost{{background:#e2e8f0;color:#334155}}
.empty{{background:#fff;padding:14px;border-radius:10px;color:#64748b}}
footer{{text-align:center;margin-top:28px;padding-top:16px;border-top:1px solid #e2e8f0;font-size:12px;color:#94a3b8}}
@media(max-width:760px){{.cards{{grid-template-columns:1fr 1fr}}.kv-grid{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<div class="wrap">
  <div class="brand-bar"></div>
  <header>
    <h1>Relatorio de Concorrentes</h1>
    <div class="sub">{_esc(nicho)} &middot; {_esc(slug)}</div>
  </header>
  <nav class="tabs">{''.join(tabs)}</nav>
  {''.join(panels)}
  <footer>Gerado a partir de {_esc(xlsx_path.name)}</footer>
</div>
<script>
document.querySelectorAll('.tab').forEach(btn => {{
  btn.addEventListener('click', () => {{
    document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('panel-' + btn.dataset.tab).classList.add('active');
  }});
}});
function applyFilters(root) {{
  const sim = (root.querySelector('.f-sim') || {{}}).value || '';
  const fonte = (root.querySelector('.f-fonte') || {{}}).value || '';
  const q = ((root.querySelector('.f-q') || {{}}).value || '').toLowerCase();
  const table = root.parentElement.querySelector('table');
  if (!table) return;
  table.querySelectorAll('tbody tr').forEach(tr => {{
    const okSim = !sim || tr.dataset.sim === sim;
    const okFonte = !fonte || tr.dataset.fonte === fonte;
    const okQ = !q || tr.innerText.toLowerCase().includes(q);
    tr.style.display = (okSim && okFonte && okQ) ? '' : 'none';
  }});
}}
document.querySelectorAll('.filters').forEach(root => {{
  root.addEventListener('input', () => applyFilters(root));
  root.addEventListener('change', () => applyFilters(root));
}});
document.querySelectorAll('.toggle-extra-cols').forEach(btn => {{
  btn.addEventListener('click', () => {{
    const root = btn.closest('.panel');
    if (!root) return;
    const show = btn.dataset.show !== '1';
    root.querySelectorAll('.extra-col').forEach(el => el.classList.toggle('is-hidden-col', !show));
    btn.dataset.show = show ? '1' : '0';
    btn.textContent = show ? 'Ocultar colunas extras' : 'Mostrar colunas extras';
  }});
}});
function cellSortValue(td) {{
  if (!td) return null;
  const raw = (td.dataset.n || '').trim();
  if (raw === '') return null;
  const n = Number(raw);
  if (Number.isFinite(n) && raw !== '' && !/[^0-9eE.+-]/.test(raw)) return n;
  return raw.toLowerCase();
}}
function sortTrafficTable(table, col, dir) {{
  const tbody = table.tBodies[0];
  const rows = Array.from(tbody.rows);
  const missing = v => v === null || v === undefined || v === '';
  rows.sort((a, b) => {{
    const av = cellSortValue(a.cells[col]);
    const bv = cellSortValue(b.cells[col]);
    if (missing(av) && missing(bv)) return 0;
    if (missing(av)) return 1;
    if (missing(bv)) return -1;
    if (typeof av === 'number' && typeof bv === 'number') {{
      return dir === 'desc' ? bv - av : av - bv;
    }}
    return dir === 'desc'
      ? String(bv).localeCompare(String(av), 'pt-BR')
      : String(av).localeCompare(String(bv), 'pt-BR');
  }});
  rows.forEach(r => tbody.appendChild(r));
}}
function sortGenericTable(table, col, dir) {{
  const tbody = table.tBodies[0];
  const rows = Array.from(tbody.rows);
  const missing = v => v === null || v === undefined || v === '';
  rows.sort((a, b) => {{
    const av = cellSortValue(a.cells[col]);
    const bv = cellSortValue(b.cells[col]);
    if (missing(av) && missing(bv)) return 0;
    if (missing(av)) return 1;
    if (missing(bv)) return -1;
    if (typeof av === 'number' && typeof bv === 'number') {{
      return dir === 'desc' ? bv - av : av - bv;
    }}
    return dir === 'desc'
      ? String(bv).localeCompare(String(av), 'pt-BR')
      : String(av).localeCompare(String(bv), 'pt-BR');
  }});
  rows.forEach(r => tbody.appendChild(r));
}}
function reorderGrowthTable(mode) {{
  const table = document.getElementById('tbl-crescimento');
  if (!table) return;
  const order = mode === 'marca'
    ? [0, 1, 2, 4, 8, 9, 10, 3, 5, 6, 7]
    : [0, 1, 2, 3, 5, 6, 7, 4, 8, 9, 10];
  table.querySelectorAll('tr').forEach(row => {{
    const cells = Array.from(row.children);
    if (cells.length < order.length) return;
    order.forEach(idx => row.appendChild(cells[idx]));
  }});
  table.querySelectorAll('th.sortable').forEach((th, idx) => {{
    th.dataset.col = String(idx);
  }});
  document.querySelectorAll('[data-growth-toggle]').forEach(btn => {{
    const active = btn.dataset.growthToggle === mode;
    btn.classList.toggle('ghost', !active);
  }});
}}
document.querySelectorAll('table.traffic th.sortable').forEach(th => {{
  th.addEventListener('click', () => {{
    const table = th.closest('table');
    const col = Number(th.dataset.col);
    const next = th.classList.contains('sorted-desc') ? 'asc' : 'desc';
    table.querySelectorAll('th.sortable').forEach(h => h.classList.remove('sorted-desc', 'sorted-asc'));
    th.classList.add(next === 'desc' ? 'sorted-desc' : 'sorted-asc');
    sortTrafficTable(table, col, next);
  }});
}});
document.querySelectorAll('table.sortable-generic th.sortable').forEach(th => {{
  th.addEventListener('click', () => {{
    const table = th.closest('table');
    const col = Number(th.dataset.col);
    const next = th.classList.contains('sorted-desc') ? 'asc' : 'desc';
    table.querySelectorAll('th.sortable').forEach(h => h.classList.remove('sorted-desc', 'sorted-asc'));
    th.classList.add(next === 'desc' ? 'sorted-desc' : 'sorted-asc');
    sortGenericTable(table, col, next);
  }});
}});
document.querySelectorAll('[data-growth-toggle]').forEach(btn => {{
  btn.addEventListener('click', () => reorderGrowthTable(btn.dataset.growthToggle));
}});
</script>
</body>
</html>"""


def generate_html_from_xlsx(xlsx_path: Path) -> Path:
    xlsx_path = Path(xlsx_path)
    print(f"[HTML] Lendo {xlsx_path.name} ...")
    sheets = load_sheets(xlsx_path)
    html_path = xlsx_path.with_suffix(".html")
    html_path.write_text(build_html(xlsx_path, sheets), encoding="utf-8")
    print(f"[HTML] Salvo: {html_path}")
    return html_path


def main():
    if len(sys.argv) < 2:
        print("Uso: python gerar_html_concorrentes.py <slug_ou_xlsx>")
        print("Ex:  python gerar_html_concorrentes.py grupodata")
        sys.exit(1)
    xlsx = find_xlsx(sys.argv[1])
    generate_html_from_xlsx(xlsx)


if __name__ == "__main__":
    main()
