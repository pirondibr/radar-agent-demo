# -*- coding: utf-8 -*-
"""Gera HTML imprimivel do relatorio de uma run (para download no admin)."""

from __future__ import annotations

import html
import json
from pathlib import Path
from typing import Any, Optional


def _esc(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def _md_lite(text: str) -> str:
    import re

    raw = _esc(text or "")
    return re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", raw)


def _rows_table(headers: list[str], rows: list[list[str]]) -> str:
    thead = "".join(f"<th>{_esc(h)}</th>" for h in headers)
    body = []
    for row in rows:
        tds = "".join(f"<td>{cell}</td>" for cell in row)
        body.append(f"<tr>{tds}</tr>")
    if not body:
        body.append(f'<tr><td colspan="{len(headers)}">Sem dados</td></tr>')
    return (
        f'<div class="table-wrap"><table><thead><tr>{thead}</tr></thead>'
        f"<tbody>{''.join(body)}</tbody></table></div>"
    )


def _section_rank(title: str, section: Optional[dict[str, Any]], kind: str = "ads") -> str:
    if not section:
        return ""
    rows_data = section.get("rows") or []
    if not rows_data and not section.get("insight"):
        return ""

    table_rows: list[list[str]] = []
    if kind == "ads":
        headers = ["#", "Empresa", "Anúncios", "Investimento", "%", "Similaridade"]
        for i, r in enumerate(rows_data, 1):
            name = _esc(r.get("name") or r.get("domain") or "—")
            if r.get("is_client"):
                name += ' <span class="you">você</span>'
            table_rows.append([
                str(i),
                name,
                _esc(r.get("ads_fmt") or r.get("value_fmt") or "0"),
                _esc(r.get("investimento_fmt") or "—"),
                _esc(r.get("pct_fmt") or "—"),
                _esc(r.get("similaridade") or "—"),
            ])
        hero = _esc(section.get("total_invest_fmt") or section.get("total_fmt") or "—")
        hero_sub = f"{_esc(section.get('total_ads_fmt') or section.get('total_fmt') or '')} · líder {_esc(section.get('leader') or '—')}"
    elif kind in ("seo", "brand"):
        headers = ["#", "Empresa", "Valor", "Crescimento", "Similaridade"]
        for i, r in enumerate(rows_data, 1):
            name = _esc(r.get("name") or r.get("domain") or "—")
            if r.get("is_client"):
                name += ' <span class="you">você</span>'
            table_rows.append([
                str(i),
                name,
                _esc(r.get("traffic_fmt") or "0"),
                _esc(r.get("growth_fmt") or "n/d"),
                _esc(r.get("similaridade") or "—"),
            ])
        hero = _esc(section.get("total_traffic_fmt") or "—")
        hero_sub = f"líder {_esc(section.get('leader') or '—')}"
    else:
        headers = ["#", "Empresa", "Valor", "Similaridade"]
        for i, r in enumerate(rows_data, 1):
            name = _esc(r.get("name") or r.get("domain") or "—")
            if r.get("is_client"):
                name += ' <span class="you">você</span>'
            table_rows.append([
                str(i),
                name,
                _esc(r.get("value_fmt") or r.get("ads_fmt") or "0"),
                _esc(r.get("similaridade") or "—"),
            ])
        hero = _esc(section.get("total_fmt") or section.get("total_invest_fmt") or "—")
        hero_sub = f"{_esc(section.get('unit') or '')} · líder {_esc(section.get('leader') or '—')}"

    insight = section.get("insight") or ""
    analysis = section.get("analysis_title") or ""
    return f"""
    <section>
      <h2>{_esc(title)}</h2>
      <div class="hero"><div class="hero-val">{hero}</div><div class="hero-sub">{hero_sub}</div></div>
      {_rows_table(headers, table_rows)}
      {f'<p class="insight"><strong>{_esc(analysis)}</strong><br>{_md_lite(insight)}</p>' if insight else ''}
    </section>
    """


def render_report_html(
    report: dict[str, Any],
    *,
    run_meta: Optional[dict[str, Any]] = None,
) -> str:
    meta = run_meta or {}
    client = report.get("client") or meta.get("company") or meta.get("slug") or "Cliente"
    briefing = report.get("briefing") or {}
    comps = report.get("competitors") or []
    note = report.get("competitors_note") or ""

    comp_rows = []
    for i, c in enumerate(comps, 1):
        name = _esc(c.get("name") or c.get("domain") or "—")
        if c.get("is_client"):
            name += ' <span class="you">você</span>'
        comp_rows.append([
            str(i),
            name,
            _esc(c.get("domain") or "—"),
            _esc(c.get("similaridade") or ("Cliente" if c.get("is_client") else "—")),
            _esc(c.get("nicho") or "—"),
        ])

    sections = [
        _section_rank("Google Ads", report.get("google_ads"), "ads"),
        _section_rank("SEO Orgânico", report.get("seo"), "seo"),
        _section_rank("Popularidade de Marca", report.get("brand"), "brand"),
        _section_rank("Meta Ads", report.get("meta"), "meta"),
        _section_rank("LinkedIn", report.get("linkedin"), "count"),
        _section_rank("Instagram", report.get("instagram"), "count"),
        _section_rank("YouTube", report.get("youtube"), "count"),
        _section_rank("TikTok", report.get("tiktok"), "count"),
    ]

    run_line = " · ".join(
        x for x in [
            _esc(meta.get("id") or ""),
            _esc(meta.get("slug") or ""),
            _esc(meta.get("url") or briefing.get("url") or ""),
            _esc(meta.get("status") or ""),
        ] if x
    )

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Radar — {_esc(client)}</title>
<style>
  :root {{ --text:#0f172a; --muted:#64748b; --border:#e2e8f0; --accent:#4f46e5; --bg:#f8fafc; }}
  * {{ box-sizing:border-box; }}
  body {{ font-family: Inter, system-ui, sans-serif; color:var(--text); background:#fff; margin:0; padding:28px 24px 60px; line-height:1.5; }}
  .wrap {{ max-width:920px; margin:0 auto; }}
  h1 {{ font-size:28px; margin:0 0 6px; letter-spacing:-.02em; }}
  h2 {{ font-size:18px; margin:0 0 12px; color:var(--accent); }}
  .sub {{ color:var(--muted); font-size:13px; margin-bottom:22px; font-family:ui-monospace,monospace; }}
  .card {{ border:1px solid var(--border); border-radius:12px; padding:18px 20px; margin-bottom:18px; background:#fff; }}
  .hero {{ display:flex; gap:10px; align-items:baseline; flex-wrap:wrap; margin-bottom:12px; }}
  .hero-val {{ font-size:28px; font-weight:800; font-family:ui-monospace,monospace; }}
  .hero-sub {{ color:var(--muted); font-size:13px; }}
  .table-wrap {{ overflow-x:auto; }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; }}
  th, td {{ text-align:left; padding:8px 10px; border-bottom:1px solid var(--border); vertical-align:top; }}
  th {{ font-size:11px; text-transform:uppercase; letter-spacing:.04em; color:var(--muted); }}
  .you {{ color:var(--accent); font-weight:700; font-size:11px; margin-left:4px; }}
  .insight {{ font-size:13px; color:#334155; background:var(--bg); border-radius:8px; padding:12px 14px; }}
  .note {{ color:#b45309; font-size:13px; margin-top:8px; }}
  .toolbar {{ position:sticky; top:0; background:rgba(255,255,255,.92); backdrop-filter:blur(6px); padding:10px 0 14px; margin-bottom:8px; display:flex; gap:8px; flex-wrap:wrap; z-index:5; }}
  .toolbar button {{ border:1px solid var(--border); background:#fff; border-radius:8px; padding:8px 12px; font:inherit; font-size:13px; font-weight:600; cursor:pointer; }}
  .toolbar button.primary {{ background:var(--accent); color:#fff; border-color:var(--accent); }}
  @media print {{
    .toolbar {{ display:none !important; }}
    body {{ padding:0; }}
    .card, section {{ break-inside:avoid; }}
  }}
</style>
</head>
<body>
<div class="wrap">
  <div class="toolbar">
    <button type="button" class="primary" onclick="window.print()">Salvar / imprimir PDF</button>
    <button type="button" onclick="window.close()">Fechar</button>
  </div>
  <h1>Radar da Concorrência — {_esc(client)}</h1>
  <div class="sub">{run_line}</div>

  <div class="card">
    <h2>Briefing</h2>
    <p><strong>{_esc(briefing.get('client') or client)}</strong>
      {(' · ' + _esc(briefing.get('nicho'))) if briefing.get('nicho') else ''}
      {(' · ' + _esc(briefing.get('escopo'))) if briefing.get('escopo') else ''}
    </p>
    <p>{_md_lite(briefing.get('summary') or '')}</p>
    {f"<p class='sub'>URL: {_esc(briefing.get('url') or meta.get('url') or '')}</p>" if (briefing.get('url') or meta.get('url')) else ''}
  </div>

  <div class="card">
    <h2>Concorrentes ({_esc(report.get('competitors_count') or max(0, len(comps)-1))})</h2>
    {f"<p class='note'>{_esc(note)}</p>" if note else ''}
    {_rows_table(['#', 'Empresa', 'Domínio', 'Similaridade', 'Nicho'], comp_rows)}
  </div>

  {''.join(f'<div class="card">{s}</div>' for s in sections if s)}
</div>
</body>
</html>
"""


def load_saved_report(run_id: str, runs_dir: Path) -> Optional[dict[str, Any]]:
    path = runs_dir / run_id / "report.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
