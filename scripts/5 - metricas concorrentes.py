# -*- coding: utf-8 -*-
"""Metricas de canais dos top concorrentes Alto.

Fluxo MVP:
1. Le o XLSX nacional de concorrentes ja gerado
2. Pega top 5 concorrentes com similaridade Alto
3. Reaproveita SEO/Trafego do XLSX nacional
4. Busca Meta Ads via ScrapingBee + GraphQL da Ad Library
5. Busca Google Ads direto na library pelo nome da empresa
6. Busca LinkedIn Ads via ScrapingBee (accountOwner na Ad Library)
7. Extrai links sociais da homepage e tenta ler seguidores
8. Gera XLSX + HTML com aba "Metricas Canais"

Uso:
    python "5 - metricas concorrentes.py" chatguru
    python "5 - metricas concorrentes.py" chatguru google_ads
    python "5 - metricas concorrentes.py" chatguru seo
    python "5 - metricas concorrentes.py" chatguru brand
    python "5 - metricas concorrentes.py" chatguru full
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import quote_plus, unquote, urlparse

import openpyxl
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

FINAL_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(FINAL_DIR))

from workspace_paths import (  # noqa: E402
    CONC_DIR,
    MVP_DIR,
    OUT_CONCORRENTES,
    OUT_METRICAS,
    setup_workspace,
)

for p in (MVP_DIR, CONC_DIR, FINAL_DIR):
    if p and p.exists() and str(p) not in sys.path:
        sys.path.insert(0, str(p))
if str(FINAL_DIR) in sys.path:
    sys.path.remove(str(FINAL_DIR))
sys.path.insert(0, str(FINAL_DIR))

setup_workspace()

from find_concorrentes import (  # noqa: E402
    find_client_xlsx,
    read_briefing_from_xlsx,
    normalize_domain,
)
from gerar_html_concorrentes import generate_html_from_xlsx  # noqa: E402

NATIONAL_DIR = OUT_CONCORRENTES
OUTPUT_DIR = OUT_METRICAS
TOP_COMPETITORS = 5
HEADERS = {"User-Agent": "Mozilla/5.0"}
SCRAPINGBEE_API_KEY = os.environ.get("SCRAPINGBEE_API_KEY", "").strip()
SCRAPINGBEE_URL = "https://app.scrapingbee.com/api/v1/"
META_DOC_ID_TYPEAHEAD = "9755915494515334"
META_DOC_ID_ADS_COUNT = "25863916039859122"
LINKEDIN_AD_LIBRARY_BASE = "https://www.linkedin.com/ad-library/search"

HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
THIN = Side(style="thin", color="CBD5E1")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_SIM = {
    "Alto": PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid"),
    "Medio": PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
    "Baixo": PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"),
}
FILL_PERFIL = {
    "Especialista": PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"),
    "Generalista": PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
}


def slugify_client(client_input: str) -> str:
    raw = client_input.strip().lower()
    raw = raw.replace("https://", "").replace("http://", "").replace("www.", "").split("/")[0]
    return re.sub(r"[^a-z0-9]+", "", raw.split(".")[0])


def fold(value: str) -> str:
    s = unicodedata.normalize("NFKD", str(value or ""))
    s = s.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().lower()


def compact(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", fold(value))


def ascii_text(value: str) -> str:
    return unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii")


def domain_token(domain: str) -> str:
    d = normalize_domain(domain)
    d = d.replace(".com.br", "").replace(".com", "").replace(".br", "")
    d = d.replace(".net", "").replace(".org", "").replace(".io", "").replace(".ai", "")
    return compact(d)


def root_url(value: str) -> str:
    raw = (value or "").strip()
    if not raw:
        return ""
    if "://" not in raw:
        raw = "https://" + raw
    parsed = urlparse(raw)
    host = parsed.netloc or parsed.path
    return f"{parsed.scheme or 'https'}://{host}/" if host else raw


def extract_website_name(url: str) -> str:
    parsed = urlparse(root_url(url))
    host = parsed.netloc or parsed.path
    if host.startswith("www."):
        host = host[4:]
    return host.split(".")[0].strip().lower()


def normalize_token(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (value or "").lower())


def social_search_label(site_url: str, company_name: str = "") -> str:
    if company_name:
        return company_name.strip()
    host = urlparse(root_url(site_url)).netloc.lower()
    if host.startswith("www."):
        host = host[4:]
    if host.endswith(".com.br"):
        host = host[:-7]
    elif host.endswith(".com"):
        host = host[:-4]
    return re.sub(r"[^a-z0-9]+", " ", host).strip()


def _extract_google_result_urls(html_text: str) -> list[str]:
    urls = []
    seen = set()
    for raw in re.findall(r'href="/url\\?q=([^"&]+)[^"]*"', html_text or ""):
        url = unquote(raw)
        dom = normalize_domain(url)
        if not dom or dom.endswith("google.com") or dom.endswith("google.com.br"):
            continue
        if url not in seen:
            seen.add(url)
            urls.append(url)
    for raw in re.findall(r'https?://[^\\s"<>]+', html_text or ""):
        if "google." in raw:
            continue
        clean = raw.rstrip(').,\'"')
        dom = normalize_domain(clean)
        if not dom:
            continue
        if clean not in seen:
            seen.add(clean)
            urls.append(clean)
    return urls


def company_query_for_competitor(comp: dict) -> str:
    title = re.split(r"[|\-:]", comp.get("title", "") or "")[0].strip()
    title_fold = compact(title)
    if title and len(title_fold) >= 4:
        return title
    token = domain_token(comp.get("domain", ""))
    return token or comp.get("domain", "")


def parse_int(value) -> int | None:
    if value in (None, "", "—", "-", "Nao consultado"):
        return None
    raw = str(value).strip().replace(".", "").replace(",", "")
    m = re.search(r"\d+", raw)
    if not m:
        return None
    try:
        return int(m.group(0))
    except ValueError:
        return None


def parse_human_number(value: str) -> int | None:
    s = str(value or "").strip().lower().replace(" ", "")
    if not s:
        return None
    m = re.search(r"(\d+(?:[.,]\d+)?)(mil|mi|k|m|b)?", s)
    if not m:
        return parse_int(value)
    raw_num = m.group(1)
    suf = m.group(2) or ""
    if not suf:
        digits = re.sub(r"\D+", "", raw_num)
        return int(digits) if digits else None
    num = float(raw_num.replace(".", "").replace(",", "."))
    mult = 1
    if suf in ("k", "mil"):
        mult = 1_000
    elif suf in ("m", "mi"):
        mult = 1_000_000
    elif suf == "b":
        mult = 1_000_000_000
    return int(num * mult)


def load_sheet_rows(xlsx_path: Path, sheet_name: str) -> list[list[str]]:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return []
    ws = wb[sheet_name]
    rows = []
    for row in ws.iter_rows(values_only=True):
        vals = []
        for cell in row:
            if cell is None:
                vals.append("")
            elif isinstance(cell, float) and cell == int(cell):
                vals.append(str(int(cell)))
            else:
                vals.append(str(cell).strip())
        rows.append(vals)
    wb.close()
    return rows


def find_national_xlsx(client_slug: str) -> Path:
    folder = NATIONAL_DIR / client_slug
    patterns = (
        f"concorrentes-all-{client_slug}-nacional.xlsx",
        f"concorrentes-all-{client_slug}*.xlsx",
    )
    for pattern in patterns:
        files = sorted(folder.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
        if files:
            return files[0]
    raise FileNotFoundError(f"XLSX nacional nao encontrado em {folder}")


def load_top_alto_competitors(xlsx_path: Path, limit: int = TOP_COMPETITORS) -> list[dict]:
    rows = load_sheet_rows(xlsx_path, "Concorrentes Unificado")
    if not rows:
        return []
    headers = rows[0]
    idx = {h: i for i, h in enumerate(headers)}

    def cell(row: list[str], name: str) -> str:
        i = idx.get(name)
        return row[i] if i is not None and i < len(row) else ""

    altos: list[dict] = []
    user_added: list[dict] = []
    for row in rows[1:]:
        if not any(row):
            continue
        if cell(row, "Similaridade") != "Alto":
            continue
        item = {
            "domain": normalize_domain(cell(row, "Dominio")),
            "url": cell(row, "URL"),
            "title": cell(row, "Titulo"),
            "similaridade": cell(row, "Similaridade"),
            "perfil": cell(row, "Perfil"),
            "fonte": cell(row, "Fonte"),
            "nicho": cell(row, "Nicho (LLM)"),
        }
        if not item["domain"]:
            continue
        fonte_l = (item["fonte"] or "").strip().lower()
        # Concorrentes do usuario (enrich) entram sempre, mesmo alem do top N
        if "usuario" in fonte_l or "user" in fonte_l:
            user_added.append(item)
        else:
            altos.append(item)

    out = altos[:limit]
    seen = {c["domain"] for c in out}
    for u in user_added:
        if u["domain"] in seen:
            continue
        out.append(u)
        seen.add(u["domain"])
    return out


def load_traffic_map(xlsx_path: Path) -> dict[str, dict]:
    rows = load_sheet_rows(xlsx_path, "Trafego SEO Marca")
    if len(rows) < 4:
        return {}
    headers = rows[2]
    idx = {h: i for i, h in enumerate(headers)}
    cur_seo = next((h for h in headers if h.startswith("SEO ")), "")
    cur_brand = next((h for h in headers if h.startswith("Marca ")), "")
    cur_total = next((h for h in headers if h.startswith("Total ")), "")
    out: dict[str, dict] = {}
    for row in rows[3:]:
        if not any(row):
            continue
        dom = normalize_domain(row[idx.get("Dominio", 1)])
        out[dom] = {
            "tipo": row[idx.get("Tipo", 0)],
            "perfil": row[idx.get("Perfil", 2)],
            "seo_atual": parse_int(row[idx.get(cur_seo, -1)]) if cur_seo in idx else None,
            "marca_atual": parse_int(row[idx.get(cur_brand, -1)]) if cur_brand in idx else None,
            "total_atual": parse_int(row[idx.get(cur_total, -1)]) if cur_total in idx else None,
            "cresc_seo_1a": row[idx.get("Cresc SEO 1a (%)", -1)] if "Cresc SEO 1a (%)" in idx else "",
            "cresc_seo_2a": row[idx.get("Cresc SEO 2a (%)", -1)] if "Cresc SEO 2a (%)" in idx else "",
            "cresc_seo_3a": row[idx.get("Cresc SEO 3a (%)", -1)] if "Cresc SEO 3a (%)" in idx else "",
            "cresc_marca_1a": row[idx.get("Cresc Marca 1a (%)", -1)] if "Cresc Marca 1a (%)" in idx else "",
            "cresc_marca_2a": row[idx.get("Cresc Marca 2a (%)", -1)] if "Cresc Marca 2a (%)" in idx else "",
            "cresc_marca_3a": row[idx.get("Cresc Marca 3a (%)", -1)] if "Cresc Marca 3a (%)" in idx else "",
        }
    return out


def build_client_entity(briefing: dict, traffic_map: dict[str, dict]) -> dict:
    dom = normalize_domain(briefing.get("url", ""))
    traffic = traffic_map.get(dom, {})
    return {
        "domain": dom,
        "url": root_url(briefing.get("url", "")),
        "title": dom,
        "similaridade": "Cliente",
        "perfil": traffic.get("perfil") or "—",
        "fonte": "Cliente",
        "nicho": briefing.get("nicho_principal") or briefing.get("nicho") or "",
    }


def _parse_meta_graphql_response(raw: str) -> dict | None:
    """Facebook GraphQL via ScrapingBee às vezes devolve 1+ linhas JSON."""
    text = (raw or "").strip()
    if not text:
        return None
    candidates = [text]
    # resposta multi-linha (comum no Ad Library)
    parts = [p.strip() for p in text.split("\n") if p.strip()]
    if len(parts) > 1:
        candidates = [parts[1], parts[0]] + parts[2:]
    for chunk in candidates:
        for attempt in (chunk, chunk.replace('\\"', '"'), chunk.replace('\\"', "")):
            try:
                data = json.loads(attempt)
                if isinstance(data, dict):
                    return data
            except Exception:
                continue
    return None


def _meta_page_id_via_scrapingbee(query: str) -> tuple[str | None, int]:
    """Resolve page_id pelo typeahead da Ad Library (doc do scraper atual)."""
    if not query:
        return None, 0
    params = {
        "api_key": SCRAPINGBEE_API_KEY,
        "url": "https://www.facebook.com/api/graphql/",
        "country_code": "br",
    }
    data = {
        "variables": json.dumps({
            "queryString": query,
            "isMobile": False,
            "country": "BR",
            "adType": "ALL",
        }, separators=(",", ":")),
        "doc_id": META_DOC_ID_TYPEAHEAD,
    }
    # Formato legado do scraper de referencia (string embutida)
    data_legacy = {
        "variables": '{"queryString":"' + query.replace('"', "") + '","isMobile":false,"country":"BR","adType":"ALL"}',
        "doc_id": META_DOC_ID_TYPEAHEAD,
    }
    credits = 0
    for payload in (data_legacy, data):
        try:
            r = requests.post(SCRAPINGBEE_URL, params=params, data=payload, timeout=60)
            credits += int(r.headers.get("Spb-cost", 1) or 1)
            if r.status_code == 401:
                raise RuntimeError("ScrapingBee sem creditos (401) no Meta typeahead")
            r.raise_for_status()
            js = _parse_meta_graphql_response(r.text)
            if not js:
                continue
            pages = (
                ((js.get("data") or {}).get("ad_library_main") or {})
                .get("typeahead_suggestions", {})
                .get("page_results")
                or []
            )
            if pages and pages[0].get("page_id"):
                return str(pages[0]["page_id"]), credits
        except RuntimeError:
            raise
        except Exception:
            continue
    return None, credits


def _meta_ads_count_via_scrapingbee(page_id: str) -> tuple[int | None, int]:
    """Conta anuncios ACTIVE da pagina na Ad Library (BR)."""
    params = {
        "api_key": SCRAPINGBEE_API_KEY,
        "url": "https://www.facebook.com/api/graphql/",
        "country_code": "br",
        "premium_proxy": "true",
        "render_js": "false",
    }
    variables = (
        '{"activeStatus":"ACTIVE","adType":"ALL","audienceTimeframe":"LAST_7_DAYS",'
        '"bylines":[],"contentLanguages":[],"countries":["BR"],"country":"BR",'
        '"deeplinkAdID":null,"excludedIDs":[],"fetchPageInfo":true,"fetchSharedDisclaimers":true,'
        '"hasDeeplinkAdID":false,"isAboutTab":false,"isAudienceTab":false,"isLandingPage":false,'
        '"isTargetedCountry":false,"location":null,"mediaType":"ALL","multiCountryFilterMode":null,'
        '"pageIDs":[],"potentialReachInput":[],"publisherPlatforms":[],"queryString":"","regions":[],'
        '"searchType":"PAGE","shouldFetchCount":true,'
        '"sortData":{"mode":"SORT_BY_TOTAL_IMPRESSIONS","direction":"DESCENDING"},'
        '"source":null,"startDate":null,"viewAllPageID":"' + page_id + '"}'
    )
    data = {"variables": variables, "doc_id": META_DOC_ID_ADS_COUNT}
    credits = 0
    for attempt in range(3):
        try:
            r = requests.post(SCRAPINGBEE_URL, params=params, data=data, timeout=90)
            credits += int(r.headers.get("Spb-cost", 1) or 1)
            if r.status_code == 401:
                raise RuntimeError("ScrapingBee sem creditos (401) no Meta ads count")
            r.raise_for_status()
            js = _parse_meta_graphql_response(r.text)
            if not js:
                time.sleep(1.5 + attempt)
                continue
            if js.get("errors"):
                time.sleep(2 + attempt)
                continue
            count = (
                ((js.get("data") or {}).get("ad_library_main") or {})
                .get("search_results_connection", {})
                .get("count")
            )
            if count is None:
                return 0, credits
            return int(count), credits
        except RuntimeError:
            raise
        except Exception:
            time.sleep(1.5 + attempt)
    return None, credits


def lookup_meta_ads(comp: dict) -> dict:
    """Meta Ads via ScrapingBee + GraphQL (mesmo fluxo do facebook_ads_meta_scraper_atual)."""
    brand = extract_website_name(comp.get("url") or comp.get("domain", ""))
    base_url = comp.get("url") or comp.get("domain", "")
    company_query = company_query_for_competitor(comp)
    socials = social_links_from_homepage(base_url, social_search_label(base_url))

    search_terms: list[str] = []
    ig_url = socials.get("instagram") or ""
    if ig_url:
        parsed = urlparse(ig_url)
        parts = [p for p in parsed.path.split("/") if p]
        if parts:
            search_terms.append(parts[0].strip("@"))
    for term in (company_query, brand, extract_website_name(base_url), comp.get("domain", "")):
        t = str(term or "").strip()
        if t and t not in search_terms:
            search_terms.append(t)

    page_id = None
    credits = 0
    for term in search_terms:
        pid, used = _meta_page_id_via_scrapingbee(term)
        credits += used
        if pid:
            page_id = pid
            break

    if not page_id:
        return {"meta_ads": None, "meta_url": "", "spb_credits": credits}

    page_url = (
        "https://www.facebook.com/ads/library/"
        f"?active_status=active&ad_type=all&country=ALL&view_all_page_id={page_id}"
        "&search_type=page&media_type=all"
    )
    count, used = _meta_ads_count_via_scrapingbee(page_id)
    credits += used
    return {
        "meta_ads": count,
        "meta_url": page_url,
        "page_id": page_id,
        "spb_credits": credits,
    }


def load_meta_map(competitors: list[dict]) -> dict[str, dict]:
    out = {}
    n = len(competitors)
    total_credits = 0
    for i, comp in enumerate(competitors, 1):
        t0 = time.time()
        print(f"[META/SB] ({i}/{n}) {comp.get('domain', '?')} ...", flush=True)
        result = lookup_meta_ads(comp)
        out[comp["domain"]] = result
        total_credits += int(result.get("spb_credits") or 0)
        print(
            f"[META/SB] ({i}/{n}) {comp.get('domain', '?')} -> "
            f"{result.get('meta_ads')} ads  page_id={result.get('page_id') or '—'}  "
            f"credits={result.get('spb_credits', 0)}  ({time.time() - t0:.1f}s)",
            flush=True,
        )
    print(f"[META/SB] Total credits ScrapingBee: {total_credits}", flush=True)
    return out


def lookup_google_ads(company_name: str) -> dict:
    domain = normalize_domain(company_name)
    if not domain:
        domain = normalize_domain(root_url(company_name))
    data = {
        "f.req": ('{"2":40,"3":{"8":[2076],"12":{"1":"' + domain + '","2":true}},"7":{"1":1,"2":0,"3":2818}}'),
    }
    params = {
        "url": "https://adstransparency.google.com/anji/_/rpc/SearchService/SearchCreatives",
        "api_key": SCRAPINGBEE_API_KEY,
        "authuser": "0",
    }
    try:
        r = requests.post("https://app.scrapingbee.com/api/v1/", params=params, data=data, timeout=60)
        r.raise_for_status()
        js = r.json()
    except Exception:
        return {"google_ads": None, "google_ads_url": ""}
    ads_count = parse_int(js.get("5"))
    profile_id = ""
    try:
        profile_id = str(js["1"][0]["1"])
    except Exception:
        profile_id = ""
    profile_url = f"https://adstransparency.google.com/advertiser/{profile_id}" if profile_id else ""
    return {"google_ads": ads_count, "google_ads_url": profile_url}


def load_google_ads_map(competitors: list[dict]) -> dict[str, dict]:
    out = {}
    n = len(competitors)
    for i, comp in enumerate(competitors, 1):
        t0 = time.time()
        print(f"[GOOGLE] ({i}/{n}) {comp.get('domain', '?')} ...", flush=True)
        out[comp["domain"]] = lookup_google_ads(comp.get("url") or comp.get("domain", ""))
        print(
            f"[GOOGLE] ({i}/{n}) {comp.get('domain', '?')} -> "
            f"{out[comp['domain']].get('google_ads')} ads  ({time.time() - t0:.1f}s)",
            flush=True,
        )
    return out


def normalize_instagram_url(url: str) -> str:
    if not url:
        return ""
    parsed = urlparse(url)
    parts = [p for p in parsed.path.split("/") if p]
    excluded = {"p", "reel", "reels", "stories", "explore", "accounts", "rsrc.php", "static", "developer", "about", "privacy", "terms"}
    if parts and parts[0].lower() not in excluded:
        return f"https://www.instagram.com/{parts[0].strip('@')}/"
    text = fetch_url_text(url, allow_proxy=True)
    soup = BeautifulSoup(text, "html.parser")
    for meta in soup.find_all("meta"):
        content = meta.get("content", "") or ""
        m = re.search(r"\(@([a-zA-Z0-9_\.]+)\)", content)
        if m and m.group(1).lower() not in excluded:
            return f"https://www.instagram.com/{m.group(1)}/"
    for match in re.findall(r'instagram\.com/([a-zA-Z0-9_\.]+)', text, flags=re.IGNORECASE):
        handle = match.strip("/").split("?")[0]
        if handle and handle.lower() not in excluded:
            return f"https://www.instagram.com/{handle}/"
    return url


def search_instagram_profile(company_name: str, site_url: str = "") -> str:
    query_base = social_search_label(site_url) if site_url else social_search_label("", company_name)
    if not query_base:
        return ""
    params = {"search": f"{query_base} instagram", "country_code": "br"}
    headers = {
        "Authorization": f"Bearer {SCRAPINGBEE_API_KEY}",
        "Spb-Accept-Language": "pt-BR,pt;q=0.9",
        "User-Agent": "Mozilla/5.0",
    }
    try:
        r = requests.get("https://app.scrapingbee.com/api/v1/google", params=params, headers=headers, timeout=60)
        r.raise_for_status()
        data = r.json()
    except Exception:
        return ""
    for key in ("organic_results", "results", "items"):
        items = data.get(key)
        if not isinstance(items, list) or not items:
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            url = (item.get("url") or item.get("link") or "").strip()
            if "instagram.com" in url.lower():
                return normalize_instagram_url(url)
    return ""


def instagram_matches_brand(instagram_url: str, company_name: str, site_url: str = "") -> bool:
    if not instagram_url:
        return False
    parsed = urlparse(instagram_url)
    parts = [p for p in parsed.path.split("/") if p]
    if not parts:
        return False
    handle = normalize_token(parts[0].strip("@"))
    company_token = normalize_token(company_name)
    domain_token = normalize_token(extract_website_name(site_url))
    candidates = [t for t in (company_token, domain_token) if t]
    return any(tok == handle or tok in handle or handle in tok for tok in candidates if len(tok) >= 3 and len(handle) >= 3)


def normalize_youtube_url(url: str) -> str:
    if not url:
        return ""
    parsed = urlparse(url)
    parts = [p for p in parsed.path.split("/") if p]
    if not parts:
        return url
    if parts[0] in {"channel", "c", "user", "@"} and len(parts) >= 2:
        return f"{parsed.scheme or 'https'}://{parsed.netloc}/{parts[0]}/{parts[1]}/"
    return f"{parsed.scheme or 'https'}://{parsed.netloc}/{parts[0]}/"


def youtube_handle_from_url(url: str) -> str:
    if not url:
        return ""
    parsed = urlparse(url)
    parts = [p for p in parsed.path.split("/") if p]
    if not parts:
        return ""
    first = parts[0].strip()
    if first.startswith("@"):
        return first.lower()
    return ""


def normalize_tiktok_url(url: str) -> str:
    if not url:
        return ""
    parsed = urlparse(url)
    parts = [p for p in parsed.path.split("/") if p]
    if parts and parts[0].startswith("@"):
        return f"{parsed.scheme or 'https'}://{parsed.netloc}/{parts[0]}/"
    return url


def social_links_from_homepage(url: str, company_name: str = "") -> dict[str, str]:
    out = {"instagram": "", "youtube": "", "tiktok": ""}
    if not url:
        return out
    html = fetch_url_text(root_url(url), allow_proxy=True)
    if not html:
        if company_name:
            out["instagram"] = search_instagram_profile(company_name, url)
        return out
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        href = a.get("href", "").strip()
        low = href.lower()
        if not out["instagram"] and "instagram.com" in low:
            out["instagram"] = normalize_instagram_url(href)
        elif not out["youtube"] and ("youtube.com" in low or "youtu.be" in low):
            out["youtube"] = normalize_youtube_url(href)
        elif not out["tiktok"] and "tiktok.com" in low:
            out["tiktok"] = normalize_tiktok_url(href)
    if company_name and (not out["instagram"] or not instagram_matches_brand(out["instagram"], company_name, url)):
        searched = search_instagram_profile(company_name, url)
        if searched:
            out["instagram"] = searched
    return out


def _fetch_via_scrapingbee(url: str, render_js: bool = False) -> str:
    if not SCRAPINGBEE_API_KEY:
        raise RuntimeError("SCRAPINGBEE_API_KEY ausente")
    params = {
        "api_key": SCRAPINGBEE_API_KEY,
        "url": url,
        "render_js": "true" if render_js else "false",
    }
    r = requests.get("https://app.scrapingbee.com/api/v1/", params=params, timeout=90)
    r.raise_for_status()
    return r.text


def fetch_url_text(
    url: str,
    allow_proxy: bool = False,
    prefer_proxy: bool = False,
    render_js: bool = False,
) -> str:
    if not url:
        return ""
    if allow_proxy and prefer_proxy:
        try:
            return _fetch_via_scrapingbee(url, render_js=render_js)
        except Exception:
            pass
    try:
        r = requests.get(url, headers=HEADERS, timeout=30)
        r.raise_for_status()
        return r.text
    except Exception:
        if not allow_proxy:
            return ""
    try:
        return _fetch_via_scrapingbee(url, render_js=render_js)
    except Exception:
        return ""


def parse_instagram_followers(url: str) -> int | None:
    # Instagram bloqueia IP direto e exige JS — ScrapingBee com render_js
    text = fetch_url_text(url, allow_proxy=True, prefer_proxy=True, render_js=True)
    if not text:
        text = fetch_url_text(url, allow_proxy=True, prefer_proxy=True, render_js=False)
    if not text:
        text = fetch_url_text(url, allow_proxy=True, prefer_proxy=False)
    if not text:
        return None
    soup = BeautifulSoup(text, "html.parser")
    for meta in soup.find_all("meta"):
        content = ascii_text(meta.get("content", "") or "")
        for pattern in (r'([0-9][0-9.,]*)\s+Followers', r'([0-9][0-9.,]*)\s+seguidores'):
            m = re.search(pattern, content, flags=re.IGNORECASE)
            if m:
                n = parse_human_number(m.group(1))
                if n is not None:
                    return n
    patterns = [
        r'"edge_followed_by"\s*:\s*\{"count"\s*:\s*(\d+)\}',
        r'content="([^"]*followers[^"]*)"',
        r'content="([^"]*seguidores[^"]*)"',
        r'([\d.,kmb]+)\s+Followers',
        r'([\d.,kmb]+)\s+seguidores',
    ]
    for pattern in patterns:
        m = re.search(pattern, text, flags=re.IGNORECASE)
        if m:
            return parse_human_number(m.group(1))
    return None


def parse_youtube_followers(url: str) -> int | None:
    text = fetch_url_text(url, allow_proxy=True, prefer_proxy=True, render_js=True)
    if not text:
        text = fetch_url_text(url, allow_proxy=True, prefer_proxy=True, render_js=False)
    if not text:
        text = fetch_url_text(url, allow_proxy=True, prefer_proxy=False)
    if not text:
        return None
    text_ascii = ascii_text(text)
    wanted_handle = youtube_handle_from_url(url)
    if wanted_handle:
        low = text_ascii.lower()
        start = 0
        while True:
            idx = low.find(wanted_handle, start)
            if idx < 0:
                break
            chunk = text_ascii[max(0, idx - 120):idx + 240]
            for pattern in (
                r'([0-9][0-9.,]*)\s*mil\s+inscritos',
                r'([0-9][0-9.,kmb]+)\s+subscribers',
                r'([0-9][0-9.,kmb]+)\s+inscritos',
            ):
                m = re.search(pattern, chunk, flags=re.IGNORECASE)
                if m:
                    token = m.group(1) + ("mil" if "mil" in pattern else "")
                    n = parse_human_number(token)
                    if n is not None:
                        return n
            start = idx + len(wanted_handle)
    soup = BeautifulSoup(text, "html.parser")
    for meta in soup.find_all("meta"):
        content = ascii_text(meta.get("content", "") or "")
        for pattern in (r'([0-9][0-9.,]*)\s*mil\s+inscritos', r'([0-9][0-9.,]*)\s+subscribers', r'([0-9][0-9.,]*)\s+inscritos'):
            m = re.search(pattern, content, flags=re.IGNORECASE)
            if m:
                token = m.group(1) + ("mil" if "mil" in pattern else "")
                n = parse_human_number(token)
                if n is not None:
                    return n
    patterns = [
        r'"subscriberCountText".*?"simpleText":"([^"]+)"',
        r'([\d.,]+mil)\s+inscritos',
        r'([\d.,kmb]+)\s+subscribers',
        r'([\d.,kmb]+)\s+inscritos',
    ]
    for pattern in patterns:
        m = re.search(pattern, text_ascii, flags=re.IGNORECASE)
        if m:
            return parse_human_number(m.group(1))
    idx = text_ascii.lower().find("inscritos")
    if idx >= 0:
        chunk = text_ascii[max(0, idx - 80):idx]
        nums = re.findall(r'([0-9][0-9.,]*)(?:\s*mil)?', chunk, flags=re.IGNORECASE)
        if nums:
            token = nums[-1]
            if "mil" in chunk.lower():
                token += "mil"
            return parse_human_number(token)
    return None


def parse_tiktok_followers(url: str) -> int | None:
    text = fetch_url_text(url, allow_proxy=True, prefer_proxy=False)
    if not text:
        return None
    text_ascii = ascii_text(text)
    soup = BeautifulSoup(text, "html.parser")
    for meta in soup.find_all("meta"):
        content = ascii_text(meta.get("content", "") or "")
        for pattern in (r'([0-9][0-9.,]*)\s*mil\s+seguidores', r'([0-9][0-9.,]*)\s+Followers', r'([0-9][0-9.,]*)\s+seguidores'):
            m = re.search(pattern, content, flags=re.IGNORECASE)
            if m:
                token = m.group(1) + ("mil" if "mil" in pattern else "")
                n = parse_human_number(token)
                if n is not None:
                    return n
    patterns = [
        r'"followerCount"\s*:\s*(\d+)',
        r'([\d.,]+mil)\s+seguidores',
        r'([\d.,kmb]+)\s+Followers',
        r'([\d.,kmb]+)\s+seguidores',
    ]
    for pattern in patterns:
        m = re.search(pattern, text_ascii, flags=re.IGNORECASE)
        if m:
            return parse_human_number(m.group(1))
    return None


def linkedin_account_owner_url(name: str) -> str:
    return f"{LINKEDIN_AD_LIBRARY_BASE}?accountOwner={quote_plus(name.strip())}&countries=BR"


def parse_linkedin_result_count(text: str) -> int | None:
    """Extrai total do h1 PT/EN/FR da LinkedIn Ad Library."""
    if not text:
        return None
    t = (text or "").replace("\xa0", " ").replace("&nbsp;", " ")
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    if re.search(r"(nenhum resultado|no results|aucun r[eé]sultat)", t, flags=re.IGNORECASE):
        return 0
    m = re.search(
        r"([\d.,\s]+)\s+(?:an[uú]ncios?|ads?|publicit[eé]s?)\s+(?:correspond|match)",
        t,
        flags=re.IGNORECASE,
    )
    if not m:
        return None
    raw = re.sub(r"[^\d]", "", m.group(1))
    try:
        return int(raw) if raw else None
    except ValueError:
        return None


def lookup_linkedin_ads(query: str) -> dict:
    """LinkedIn Ads via ScrapingBee (busca accountOwner na Ad Library)."""
    q = (query or "").strip()
    url = linkedin_account_owner_url(q) if q else ""
    if not q:
        return {
            "linkedin_ads": None,
            "linkedin_payer": "",
            "linkedin_search_url": "",
            "spb_credits": 0,
        }

    params = {
        "api_key": SCRAPINGBEE_API_KEY,
        "url": url,
        "render_js": "true",
        "premium_proxy": "true",
        "country_code": "br",
        "wait": "5000",
        "extract_rules": json.dumps({"h1": "h1"}),
    }
    credits = 0
    last_err: Exception | None = None
    for attempt in range(3):
        try:
            r = requests.get(SCRAPINGBEE_URL, params=params, timeout=120)
            credits += int(r.headers.get("Spb-cost") or 0)
            if r.status_code == 401:
                raise RuntimeError("ScrapingBee sem creditos (401) no LinkedIn")
            if r.status_code != 200:
                raise RuntimeError(f"ScrapingBee LinkedIn HTTP {r.status_code}")

            h1 = ""
            try:
                payload = r.json()
                if isinstance(payload, dict):
                    h1 = str(payload.get("h1") or "")
            except Exception:
                h1 = ""
            if not h1:
                m = re.search(r"<h1[^>]*>(.*?)</h1>", r.text or "", flags=re.IGNORECASE | re.DOTALL)
                h1 = m.group(1) if m else (r.text or "")

            count = parse_linkedin_result_count(h1)
            if count is None:
                count = parse_linkedin_result_count(r.text or "")
            return {
                "linkedin_ads": count,
                "linkedin_payer": q,
                "linkedin_search_url": url,
                "spb_credits": credits,
            }
        except Exception as exc:
            last_err = exc
            time.sleep(1.5 + attempt)

    print(f"[LINKEDIN/SB] falha '{q}': {last_err}", flush=True)
    return {
        "linkedin_ads": None,
        "linkedin_payer": q,
        "linkedin_search_url": url,
        "spb_credits": credits,
    }


def load_linkedin_map(competitors: list[dict]) -> dict[str, dict]:
    if not competitors:
        return {}
    out = {}
    n = len(competitors)
    total_credits = 0
    for i, comp in enumerate(competitors, 1):
        query = company_query_for_competitor(comp)
        t0 = time.time()
        print(f"[LINKEDIN/SB] ({i}/{n}) {comp.get('domain', '?')} query='{query}' ...", flush=True)
        result = lookup_linkedin_ads(query)
        out[comp["domain"]] = result
        total_credits += int(result.get("spb_credits") or 0)
        print(
            f"[LINKEDIN/SB] ({i}/{n}) {comp.get('domain', '?')} -> "
            f"{result.get('linkedin_ads')} ads  credits={result.get('spb_credits', 0)}  "
            f"({time.time() - t0:.1f}s)",
            flush=True,
        )
    print(f"[LINKEDIN/SB] Total credits ScrapingBee: {total_credits}", flush=True)
    return out


def build_metric_rows(
    competitors: list[dict],
    traffic_map: dict[str, dict],
    meta_map: dict[str, dict],
    google_ads_map: dict[str, dict],
    linkedin_map: dict[str, dict],
    collect_social: bool = True,
) -> list[dict]:
    rows = []
    n = len(competitors)
    for i, comp in enumerate(competitors, 1):
        t0 = time.time()
        dom = comp["domain"]
        base_url = comp.get("url") or comp.get("domain", "")
        if collect_social:
            print(f"[SOCIAL] ({i}/{n}) {dom} ...", flush=True)
            socials = social_links_from_homepage(base_url, social_search_label(base_url))
            t_links = time.time()
            instagram_followers = parse_instagram_followers(socials["instagram"]) if socials["instagram"] else None
            youtube_followers = parse_youtube_followers(socials["youtube"]) if socials["youtube"] else None
            # TikTok: script ainda nao existe — nao coletamos no chatbot
            socials["tiktok"] = ""
            tiktok_followers = None
            print(
                f"[SOCIAL] ({i}/{n}) {dom} links={time.time() - t0:.1f}s "
                f"followers_extra={time.time() - t_links:.1f}s total={time.time() - t0:.1f}s "
                f"(tiktok skipped)",
                flush=True,
            )
        else:
            socials = {"instagram": "", "youtube": "", "tiktok": ""}
            instagram_followers = youtube_followers = tiktok_followers = None
        traffic = traffic_map.get(dom, {})
        meta = meta_map.get(dom, {})
        google = google_ads_map.get(dom, {})
        linkedin = linkedin_map.get(dom, {})
        rows.append({
            "domain": dom,
            "url": comp.get("url", ""),
            "similaridade": comp.get("similaridade", "Alto"),
            "perfil": comp.get("perfil", "—"),
            "fonte": comp.get("fonte", ""),
            "seo_atual": traffic.get("seo_atual"),
            "marca_atual": traffic.get("marca_atual"),
            "meta_ads": meta.get("meta_ads"),
            "meta_ads_url": meta.get("meta_url", ""),
            "google_ads": google.get("google_ads"),
            "google_ads_url": google.get("google_ads_url", ""),
            "linkedin_ads": linkedin.get("linkedin_ads"),
            "linkedin_ads_url": linkedin.get("linkedin_search_url", ""),
            "instagram_url": socials["instagram"],
            "youtube_url": socials["youtube"],
            "tiktok_url": socials["tiktok"],
            "instagram_followers": instagram_followers,
            "youtube_followers": youtube_followers,
            "tiktok_followers": tiktok_followers,
        })
    return rows


def load_existing_channel_map(xlsx_path: Path) -> dict[str, dict]:
    """Reaproveita Google/Meta/LinkedIn ja gravados em metricas anteriores."""
    if not xlsx_path or not Path(xlsx_path).exists():
        return {}
    rows = load_sheet_rows(xlsx_path, "Metricas Canais")
    if not rows:
        return {}
    headers = rows[0]
    idx = {str(h).strip(): i for i, h in enumerate(headers)}
    out: dict[str, dict] = {}
    for row in rows[1:]:
        if not any(row):
            continue
        i_dom = idx.get("Dominio")
        if i_dom is None or i_dom >= len(row):
            continue
        dom = normalize_domain(row[i_dom] or "")
        if not dom:
            continue

        def col(*names):
            for name in names:
                j = idx.get(name)
                if j is not None and j < len(row):
                    return row[j]
            return None

        out[dom] = {
            "google_ads": col("Google Ads"),
            "google_ads_url": col("Google Ads URL") or "",
            "meta_ads": col("Meta Ads"),
            "meta_url": col("Meta Ads URL") or "",
            "linkedin_ads": col("Linkedin Ads", "LinkedIn Ads"),
            "linkedin_search_url": col("Linkedin Ads URL", "LinkedIn Ads URL") or "",
            "instagram_url": col("Instagram URL") or "",
            "youtube_url": col("YouTube URL", "Youtube URL") or "",
            "tiktok_url": col("TikTok URL", "Tiktok URL") or "",
            "instagram_followers": col("Instagram Seguidores", "Instagram Followers", "Instagram"),
            "youtube_followers": col("Youtube Seguidores", "YouTube Seguidores", "YouTube Followers", "Youtube Followers", "YouTube"),
            "tiktok_followers": col("Tiktok Seguidores", "TikTok Followers", "Tiktok Followers"),
        }
    return out


def _merge_preserved_channels(
    metric_rows: list[dict],
    existing_maps: dict[str, dict],
    *,
    keep_meta: bool = False,
    keep_google: bool = False,
    keep_linkedin: bool = False,
    keep_social: bool = False,
) -> None:
    """Preserva canais ja coletados quando o modo atual nao os recalcula."""
    for row in metric_rows:
        prev = existing_maps.get(row.get("domain") or "", {})
        if not prev:
            continue
        if keep_meta and row.get("meta_ads") is None:
            row["meta_ads"] = prev.get("meta_ads")
            row["meta_ads_url"] = prev.get("meta_url") or row.get("meta_ads_url") or ""
        if keep_google and row.get("google_ads") is None:
            row["google_ads"] = prev.get("google_ads")
            row["google_ads_url"] = prev.get("google_ads_url") or row.get("google_ads_url") or ""
        if keep_linkedin and row.get("linkedin_ads") is None:
            row["linkedin_ads"] = prev.get("linkedin_ads")
            row["linkedin_ads_url"] = prev.get("linkedin_search_url") or row.get("linkedin_ads_url") or ""
        if keep_social:
            if not row.get("instagram_url") and prev.get("instagram_url"):
                row["instagram_url"] = prev.get("instagram_url") or ""
                row["instagram_followers"] = prev.get("instagram_followers")
            if not row.get("youtube_url") and prev.get("youtube_url"):
                row["youtube_url"] = prev.get("youtube_url") or ""
                row["youtube_followers"] = prev.get("youtube_followers")
            # TikTok: nunca coletamos no chatbot (script ainda nao existe)
            if prev.get("tiktok_url"):
                row["tiktok_url"] = prev.get("tiktok_url") or ""
                row["tiktok_followers"] = prev.get("tiktok_followers")


def filtered_sheet(rows: list[list[str]], keep_domains: set[str], include_client: bool = False) -> list[list[str]]:
    if not rows:
        return []
    out = rows[:3]
    for row in rows[3:]:
        if not any(row):
            continue
        dom = normalize_domain(row[1] if len(row) > 1 else "")
        role = fold(row[0] if row else "")
        if dom in keep_domains or (include_client and role.startswith("cliente")):
            out.append(row)
    return out


def write_rows_sheet(wb: Workbook, sheet_name: str, rows: list[list[str]]) -> None:
    if not rows:
        return
    ws = wb.create_sheet(sheet_name)
    for row in rows:
        ws.append(row)
    for row_idx, row in enumerate(ws.iter_rows(), 1):
        for cell in row:
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        if row_idx in (1, 3):
            for cell in row:
                if cell.value not in (None, ""):
                    cell.font = HEAD_FONT if row_idx == 3 else Font(bold=True, size=12, color="0F172A")
                    if row_idx == 3:
                        cell.fill = HEAD_FILL


def write_metrics_xlsx(
    out_path: Path,
    client_input: str,
    briefing: dict,
    competitors: list[dict],
    metric_rows: list[dict],
    traffic_rows: list[list[str]],
    growth_rows: list[list[str]],
    national_xlsx: Path,
) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"
    ws.append(["Metricas Concorrentes"])
    ws.merge_cells("A1:B1")
    ws["A1"].font = Font(bold=True, size=14, color="0F172A")
    ws.append([])
    ws.append(["Cliente (input)", client_input])
    ws.append(["URL cliente", briefing.get("url", "")])
    ws.append(["Nicho principal", briefing.get("nicho_principal", "")])
    ws.append(["Escopo", briefing.get("escopo", "")])
    ws.append(["Top concorrentes", str(len(competitors))])
    ws.append(["Fonte nacional", str(national_xlsx)])
    ws.append(["Meta Ads", "ScrapingBee + GraphQL Ad Library (typeahead page_id + count ACTIVE)"])
    ws.append(["Google Ads", "Busca direta por empresa na Google Ads Transparency"])
    ws.append(["LinkedIn Ads", "ScrapingBee + LinkedIn Ad Library (accountOwner + h1 total)"])
    ws.append(["Data", datetime.now().strftime("%d/%m/%Y %H:%M")])
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 90

    ws_u = wb.create_sheet("Concorrentes Unificado")
    headers_u = ["#", "Dominio", "Similaridade", "Perfil", "Fonte", "Nicho (LLM)", "URL", "Titulo"]
    ws_u.append(headers_u)
    for c in ws_u[1]:
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BORDER
    for i, comp in enumerate(competitors, 1):
        ws_u.append([i, comp.get("domain", ""), comp.get("similaridade", ""), comp.get("perfil", ""), comp.get("fonte", ""), comp.get("nicho", ""), comp.get("url", ""), comp.get("title", "")])
        last = ws_u.max_row
        ws_u.cell(row=last, column=3).fill = FILL_SIM.get(comp.get("similaridade", ""), PatternFill(fill_type=None))
        ws_u.cell(row=last, column=4).fill = FILL_PERFIL.get(comp.get("perfil", ""), PatternFill(fill_type=None))
        for cell in ws_u[last]:
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws_m = wb.create_sheet("Metricas Canais")
    headers_m = [
        "#", "Dominio", "Similaridade", "Perfil", "Fonte",
        "SEO Atual", "Marca Atual", "Meta Ads", "Google Ads", "Linkedin Ads",
        "Meta Ads URL", "Google Ads URL", "Linkedin Ads URL",
        "Instagram Seguidores", "Youtube Seguidores", "Tiktok Seguidores",
        "Instagram URL", "Youtube URL", "Tiktok URL", "URL",
    ]
    ws_m.append(headers_m)
    for c in ws_m[1]:
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for i, row in enumerate(metric_rows, 1):
        ws_m.append([
            i, row["domain"], row["similaridade"], row["perfil"], row["fonte"],
            row["seo_atual"], row["marca_atual"], row["meta_ads"], row["google_ads"], row["linkedin_ads"],
            row["meta_ads_url"], row["google_ads_url"], row["linkedin_ads_url"],
            row["instagram_followers"], row["youtube_followers"], row["tiktok_followers"],
            row["instagram_url"], row["youtube_url"], row["tiktok_url"], row["url"],
        ])
        last = ws_m.max_row
        ws_m.cell(row=last, column=3).fill = FILL_SIM.get(row["similaridade"], PatternFill(fill_type=None))
        ws_m.cell(row=last, column=4).fill = FILL_PERFIL.get(row["perfil"], PatternFill(fill_type=None))
        for metric_col, url_col in ((8, 11), (9, 12), (10, 13), (14, 17), (15, 18), (16, 19)):
            metric_cell = ws_m.cell(row=last, column=metric_col)
            url_value = ws_m.cell(row=last, column=url_col).value
            if url_value not in (None, "", "—") and metric_cell.value not in (None, "", "—"):
                metric_cell.hyperlink = str(url_value)
                metric_cell.style = "Hyperlink"
        for cell in ws_m[last]:
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    widths = [4, 26, 13, 13, 14, 12, 12, 10, 10, 11, 34, 34, 34, 18, 18, 18, 34, 34, 34, 42]
    for i, w in enumerate(widths, 1):
        ws_m.column_dimensions[chr(64 + i)].width = w
    for col in ("K", "L", "M", "Q", "R", "S", "T"):
        ws_m.column_dimensions[col].hidden = True

    write_rows_sheet(wb, "Trafego SEO Marca", traffic_rows)
    write_rows_sheet(wb, "Crescimento", growth_rows)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def main() -> None:
    if len(sys.argv) < 2:
        print('Uso: python "5 - metricas concorrentes.py" <cliente> [google_ads|seo|brand|meta|linkedin|social|full]')
        sys.exit(1)

    t_all = time.time()
    timings: list[tuple[str, float]] = []

    def mark(label: str, t0: float) -> None:
        dt = time.time() - t0
        timings.append((label, dt))
        print(f"[TIME-SUB] {label}: {dt:.1f}s", flush=True)

    client_input = sys.argv[1].strip()
    mode = (sys.argv[2].strip().lower() if len(sys.argv) >= 3 else "full")
    valid_modes = {"full", "google_ads", "seo", "brand", "meta", "linkedin", "social"}
    if mode not in valid_modes:
        print(f"ERRO: modo '{mode}' invalido. Opcoes: {', '.join(sorted(valid_modes))}")
        sys.exit(1)

    client_slug = slugify_client(client_input)
    print(f"\n=== Metricas Final 09: '{client_input}' | modo {mode} ===\n", flush=True)

    t0 = time.time()
    briefing_xlsx = find_client_xlsx(client_input)
    briefing = read_briefing_from_xlsx(briefing_xlsx)
    national_xlsx = find_national_xlsx(client_slug)
    traffic_map = load_traffic_map(national_xlsx)
    competitors = load_top_alto_competitors(national_xlsx, TOP_COMPETITORS)
    if not competitors:
        raise RuntimeError("Nenhum concorrente Alto encontrado no XLSX nacional.")
    metric_entities = [build_client_entity(briefing, traffic_map)] + competitors
    mark(f"setup (briefing+top {len(competitors)} Alto)", t0)
    print(f"[METRICAS] Entidades: {len(metric_entities)} (cliente + {len(competitors)} Alto)", flush=True)

    out_path = OUTPUT_DIR / client_slug / f"metricas-concorrentes-{client_slug}.xlsx"
    existing_maps = load_existing_channel_map(out_path) if out_path.exists() else {}

    meta_map: dict[str, dict] = {}
    google_ads_map: dict[str, dict] = {}
    linkedin_map: dict[str, dict] = {}
    collect_social = False
    keep_meta = keep_google = keep_linkedin = keep_social = False

    if mode == "full":
        t0 = time.time()
        print("[METRICAS] Subetapa Meta Ads Library (ScrapingBee GraphQL) ...", flush=True)
        meta_map = load_meta_map(metric_entities)
        mark("Meta Ads (ScrapingBee)", t0)

        t0 = time.time()
        print("[METRICAS] Subetapa Google Ads Transparency (ScrapingBee) ...", flush=True)
        google_ads_map = load_google_ads_map(metric_entities)
        mark("Google Ads (ScrapingBee)", t0)

        t0 = time.time()
        print("[METRICAS] Subetapa LinkedIn Ad Library (ScrapingBee) ...", flush=True)
        linkedin_map = load_linkedin_map(metric_entities)
        mark("LinkedIn Ads (ScrapingBee)", t0)
        collect_social = True

    elif mode == "google_ads":
        t0 = time.time()
        print("[METRICAS] Subetapa Google Ads Transparency (somente este canal) ...", flush=True)
        google_ads_map = load_google_ads_map(metric_entities)
        mark("Google Ads (ScrapingBee)", t0)
        keep_meta = keep_linkedin = keep_social = True
        for dom, prev in existing_maps.items():
            meta_map[dom] = {"meta_ads": prev.get("meta_ads"), "meta_url": prev.get("meta_url", "")}
            linkedin_map[dom] = {
                "linkedin_ads": prev.get("linkedin_ads"),
                "linkedin_search_url": prev.get("linkedin_search_url", ""),
            }

    elif mode == "meta":
        t0 = time.time()
        print("[METRICAS] Subetapa Meta Ads Library (somente este canal) ...", flush=True)
        meta_map = load_meta_map(metric_entities)
        mark("Meta Ads (ScrapingBee)", t0)
        keep_google = keep_linkedin = keep_social = True
        for dom, prev in existing_maps.items():
            google_ads_map[dom] = {
                "google_ads": prev.get("google_ads"),
                "google_ads_url": prev.get("google_ads_url", ""),
            }
            linkedin_map[dom] = {
                "linkedin_ads": prev.get("linkedin_ads"),
                "linkedin_search_url": prev.get("linkedin_search_url", ""),
            }

    elif mode == "linkedin":
        t0 = time.time()
        print("[METRICAS] Subetapa LinkedIn Ads (somente este canal) ...", flush=True)
        linkedin_map = load_linkedin_map(metric_entities)
        mark("LinkedIn Ads (ScrapingBee)", t0)
        keep_meta = keep_google = keep_social = True
        for dom, prev in existing_maps.items():
            google_ads_map[dom] = {
                "google_ads": prev.get("google_ads"),
                "google_ads_url": prev.get("google_ads_url", ""),
            }
            meta_map[dom] = {"meta_ads": prev.get("meta_ads"), "meta_url": prev.get("meta_url", "")}

    elif mode == "social":
        t0 = time.time()
        print("[METRICAS] Subetapa Instagram + YouTube (sem TikTok) ...", flush=True)
        collect_social = True
        keep_meta = keep_google = keep_linkedin = True
        for dom, prev in existing_maps.items():
            google_ads_map[dom] = {
                "google_ads": prev.get("google_ads"),
                "google_ads_url": prev.get("google_ads_url", ""),
            }
            meta_map[dom] = {"meta_ads": prev.get("meta_ads"), "meta_url": prev.get("meta_url", "")}
            linkedin_map[dom] = {
                "linkedin_ads": prev.get("linkedin_ads"),
                "linkedin_search_url": prev.get("linkedin_search_url", ""),
            }
        mark("setup maps preservados (social)", t0)

    elif mode in ("seo", "brand"):
        print(
            f"[METRICAS] Modo {mode}: reaproveitando SEO/Marca do XLSX nacional (sem ScrapingBee).",
            flush=True,
        )
        keep_meta = keep_google = keep_linkedin = keep_social = True
        for dom, prev in existing_maps.items():
            google_ads_map[dom] = {
                "google_ads": prev.get("google_ads"),
                "google_ads_url": prev.get("google_ads_url", ""),
            }
            meta_map[dom] = {"meta_ads": prev.get("meta_ads"), "meta_url": prev.get("meta_url", "")}
            linkedin_map[dom] = {
                "linkedin_ads": prev.get("linkedin_ads"),
                "linkedin_search_url": prev.get("linkedin_search_url", ""),
            }

    if mode == "full":
        t0 = time.time()
        print("[METRICAS] Subetapa redes sociais (homepage + followers, sem TikTok) ...", flush=True)
        metric_rows = build_metric_rows(
            metric_entities, traffic_map, meta_map, google_ads_map, linkedin_map, collect_social=True,
        )
        mark("Redes sociais (links+followers)", t0)
    else:
        t0 = time.time()
        metric_rows = build_metric_rows(
            metric_entities, traffic_map, meta_map, google_ads_map, linkedin_map, collect_social=collect_social,
        )
        _merge_preserved_channels(
            metric_rows,
            existing_maps,
            keep_meta=keep_meta,
            keep_google=keep_google,
            keep_linkedin=keep_linkedin,
            keep_social=keep_social and not collect_social,
        )
        mark(f"Montagem Metricas Canais ({mode})", t0)

    t0 = time.time()
    keep_domains = {c["domain"] for c in competitors}
    traffic_rows = filtered_sheet(load_sheet_rows(national_xlsx, "Trafego SEO Marca"), keep_domains, include_client=True)
    growth_rows = filtered_sheet(load_sheet_rows(national_xlsx, "Crescimento"), keep_domains, include_client=True)
    saved = write_metrics_xlsx(
        out_path,
        client_input,
        briefing,
        competitors,
        metric_rows,
        traffic_rows,
        growth_rows,
        national_xlsx,
    )
    # HTML completo so no modo full (evita custo nas etapas separadas do chat)
    if mode == "full":
        html_path = generate_html_from_xlsx(saved)
        mark("Gravacao XLSX+HTML", t0)
        print(f"[OUT] HTML salvo: {html_path}")
    else:
        mark(f"Gravacao XLSX ({mode})", t0)

    print(f"[OUT] XLSX salvo: {saved}")
    print("\n[METRICAS] Resumo de subetapas:", flush=True)
    for label, dt in timings:
        print(f"  - {label}: {dt:.1f}s", flush=True)
    print(f"  - TOTAL: {time.time() - t_all:.1f}s", flush=True)


if __name__ == "__main__":
    main()
