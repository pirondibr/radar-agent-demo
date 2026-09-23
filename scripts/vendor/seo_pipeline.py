# -*- coding: utf-8 -*-
"""
SEO Pipeline padronizado.

Entrada: URL do site da empresa
Etapa 1: Scraping + LLM = briefing padronizado (14 perguntas)
Etapa 2: LLM gera keywords (Local: 20 / Nacional: 50 em topo+fundo de funil)
Etapa 3: DataForSEO Google Ads search volume (volume mensal, 1 ano atras, crescimento)
Saidas:
  - XLSX: aba Briefing + aba Palavras-chave
  - HTML: relatorio visual

Uso:
    python seo_pipeline.py <url>
    python seo_pipeline.py https://cnempreendimentos.net/
"""

import os
import sys
import re
import json
import time
import random
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path
from string import ascii_lowercase, digits
from urllib.parse import urlparse
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "").strip()
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
OPENROUTER_MODEL = "google/gemini-2.5-flash"

DATAFORSEO_USER = (
    os.environ.get("DATAFORSEO_USER", "").strip()
    or os.environ.get("DATAFORSEO_LOGIN", "").strip()
)
DATAFORSEO_PASS = (
    os.environ.get("DATAFORSEO_PASS", "").strip()
    or os.environ.get("DATAFORSEO_PASSWORD", "").strip()
)
DATAFORSEO_VOLUME_URL = "https://api.dataforseo.com/v3/keywords_data/google_ads/search_volume/live"
DATAFORSEO_SERP_URL = "https://api.dataforseo.com/v3/serp/google/organic/live/advanced"
LOCATION_BRAZIL = 2076
SERP_DEPTH = 20

SEMRUSH_API_KEY = os.environ.get("SEMRUSH_API_KEY", "").strip()
SEMRUSH_RPC_URL = "https://www.semrush.com/dpa/rpc"
SEMRUSH_DB = "br"
SEMRUSH_TOP_N = 30

MAX_LOCAL = 20
MAX_NACIONAL = 50
OVERGEN_FACTOR = 1.5

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"


# ---------------------------------------------------------------------------
# UTIL
# ---------------------------------------------------------------------------

def slugify(value: str) -> str:
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = re.sub(r"[^a-zA-Z0-9]+", "-", value).strip("-").lower()
    return value or "site"


def domain_slug(url: str) -> str:
    netloc = urlparse(url).netloc or url
    netloc = netloc.replace("www.", "")
    return slugify(netloc.split(".")[0])


def fmt_int(v):
    if v is None:
        return "—"
    try:
        return f"{int(v):,}".replace(",", ".")
    except Exception:
        return str(v)


def fmt_pct(v):
    if v is None:
        return "—"
    sign = "+" if v > 0 else ""
    return f"{sign}{v:.1f}%"


# ---------------------------------------------------------------------------
# ETAPA 1A: SCRAPING
# ---------------------------------------------------------------------------

def scrape_site(url: str, max_chars: int = 12000) -> dict:
    print(f"[1A] Scraping {url} ...")
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    }
    out = {"title": "", "meta_description": "", "h1": [], "h2": [], "nav": [], "body": ""}
    try:
        r = requests.get(url, headers=headers, timeout=30, allow_redirects=True)
        r.raise_for_status()
    except Exception as e:
        print(f"[1A] Erro no scraping: {e}")
        return out

    soup = BeautifulSoup(r.text, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    if soup.title and soup.title.string:
        out["title"] = soup.title.string.strip()
    m = soup.find("meta", attrs={"name": "description"})
    if m and m.get("content"):
        out["meta_description"] = m["content"].strip()

    out["h1"] = [h.get_text(strip=True) for h in soup.find_all("h1") if h.get_text(strip=True)][:10]
    out["h2"] = [h.get_text(strip=True) for h in soup.find_all("h2") if h.get_text(strip=True)][:20]

    nav_items = []
    for nav in soup.find_all(["nav", "header"]):
        for a in nav.find_all("a"):
            txt = a.get_text(strip=True)
            if txt and 2 < len(txt) < 40:
                nav_items.append(txt)
    out["nav"] = list(dict.fromkeys(nav_items))[:30]

    text = soup.get_text(separator=" ", strip=True)
    text = re.sub(r"\s+", " ", text)
    out["body"] = text[:max_chars]
    return out


def site_signals_text(site: dict) -> str:
    return (
        f"TITLE: {site.get('title','')}\n"
        f"META: {site.get('meta_description','')}\n"
        f"H1: {' | '.join(site.get('h1',[]))}\n"
        f"H2: {' | '.join(site.get('h2',[]))}\n"
        f"NAV: {' | '.join(site.get('nav',[]))}\n"
        f"CONTENT:\n{site.get('body','')}"
    )


# ---------------------------------------------------------------------------
# ETAPA 1B: LLM BRIEFING
# ---------------------------------------------------------------------------

BRIEFING_FIELDS = [
    ("nicho", "Nicho"),
    ("publico_alvo", "Público alvo"),
    ("audiencia_alvos", "Exemplo de Audiência (lista direta de alvos)"),
    ("b2b_b2c", "B2B ou B2C"),
    ("escopo", "Local ou Nacional"),
    ("tipo_negocio", "Serviço, Digital, Ecommerce"),
    ("modelo_produto", "Modelo de Produto"),
    ("dores", "Dores que resolve"),
    ("ferramenta_saas", "Ferramenta/SaaS"),
    ("curso", "Curso"),
    ("ai", "AI"),
    ("nicho_principal", "Nicho principal (máximo 3 palavras)"),
    ("nicho_secundario", "Nicho secundário (2 palavras)"),
    ("seo_keywords", "Defina 2 palavras para SEO"),
]


def call_openrouter(prompt: str, max_tokens: int = 1500, temperature: float = 0.2) -> str:
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": max_tokens,
        "temperature": temperature,
    }
    r = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=90)
    r.raise_for_status()
    data = r.json()
    return data["choices"][0]["message"]["content"]


def extract_json(text: str) -> dict:
    text = text.strip()
    m = re.search(r"```(?:json)?\s*(.*?)```", text, re.DOTALL)
    if m:
        text = m.group(1).strip()
    text = re.sub(r"^[^{\[]*", "", text)
    text = re.sub(r"[^}\]]*$", "", text)
    return json.loads(text)


def llm_briefing(url: str, site: dict) -> dict:
    print("[1B] Gerando briefing com LLM ...")
    fields_str = "\n".join(f'  - "{k}": {label}' for k, label in BRIEFING_FIELDS)
    prompt = f"""Voce e um analista de SEO. Leia o conteudo do site abaixo e responda em JSON valido com cada campo preenchido.

URL: {url}

CONTEUDO DO SITE:
{site_signals_text(site)}

RETORNE APENAS JSON com as chaves a seguir (use string em todas):
{fields_str}

REGRAS:
- "escopo": apenas "Local" ou "Nacional" (Local se o site atende uma regiao/cidade especifica; Nacional se atende todo o Brasil).
- Se "escopo" = "Local", adicione tambem o campo "cidades": lista de 1 a 4 cidades/regioes alvo (em portugues, sem acento, ex: ["itajai", "balneario camboriu"]).
- Se "escopo" = "Nacional", adicione "cidades": [].
- "nicho_principal": no maximo 3 palavras. USE O TERMO MAIS GENERICO E DE MAIOR BUSCA do segmento (ex: "construtora", "clinica medica", "plano saude pet"). NAO inclua adjetivos qualificadores como "luxo", "alto padrao", "premium", "exclusivo" a nao ser que apareçam claramente no TITLE/META/H1/NAV do site.
- "nicho_secundario": exatamente 2 palavras, tambem termo generico.
- "seo_keywords": string com 2 termos genericos separados por ", " (ex: "construtora itajai, apartamento praia brava").
- "ferramenta_saas", "curso", "ai": responda "Sim" ou "Nao" (com justificativa curta se preciso).
- "modelo_produto": classifica A FORMA do que o player entrega, em DOIS NIVEIS (igual a Setor -> Nicho).
   Formato OBRIGATORIO: "<Categoria 1> -> <Categoria 2>" (use a seta " -> ", sem aspas, sem texto extra).
   Categoria 1 = "Digital" | "Fisico" | "Servico".
   Categoria 2 deve ser EXATAMENTE uma das opcoes abaixo, conforme a Categoria 1 escolhida:
     * Se Categoria 1 = "Digital": "SaaS / assinatura" | "Plataforma / Marketplace" | "Infra / API / CPaaS" | "Midia / conteudo digital" | "Infoproduto / curso" | "App de consumo"
     * Se Categoria 1 = "Fisico":  "Industria / fabricante" | "E-commerce / varejo online" | "Varejo fisico / loja" | "Distribuicao / atacado"
     * Se Categoria 1 = "Servico": "Consultoria / assessoria" | "Agencia" | "Educacao / treinamento" | "Eventos / feiras" | "Associacao / conselho" | "Servico financeiro"

   *** CRITERIO DECISIVO ***
   Pergunte-se: "O que o cliente final RECEBE como entrega principal?" — NAO se baseie no tom comercial do site.
     * Se recebe um BEM TANGIVEL fabricado/produzido/construido pela empresa  -> "Fisico -> Industria / fabricante".
       Inclui: construtora, incorporadora, fabrica, industria, montadora, refinaria, agroindustria,
       empresa que tem fabrica/canteiro de obras e vende o que produz. Mesmo que tambem faca venda direta,
       a entrega principal e fisica/manufaturada.
     * Se recebe um BEM TANGIVEL revendido (nao fabricado pela empresa) por loja online -> "Fisico -> E-commerce / varejo online".
     * Se recebe um BEM TANGIVEL revendido em loja fisica -> "Fisico -> Varejo fisico / loja".
     * Se recebe BEM TANGIVEL em grande volume para outros vendedores -> "Fisico -> Distribuicao / atacado".
     * Se recebe ACESSO a um software/plataforma/API (sem produto fisico) -> categoria "Digital".
     * Se recebe HORAS DE TRABALHO/EXPERTISE humana (sem produto proprio fisico nem software proprio) -> categoria "Servico".

   *** ANTI-REGRAS (NAO USE) ***
   - NAO use "Servico -> Consultoria / assessoria" para empresas que vendem PRODUTO PROPRIO (construtoras,
     industrias, fabricantes, lojas, e-commerces). "Consultoria/assessoria" e RESERVADO para empresas cuja
     entrega central e tempo/expertise humana (ex: consultoria estrategica, assessoria juridica, contabil,
     financeira, de imprensa). Uma construtora que "assessora a compra" continua entregando um imovel fisico.
   - NAO use "Servico -> Agencia" para empresas que vendem produto proprio. "Agencia" e reservado para
     prestadores que entregam servicos de comunicacao/marketing/publicidade para terceiros.
   - NAO use "Digital" para empresas que so tem SITE/INSTAGRAM mas vendem produto fisico ou servico humano.

   *** EXEMPLOS DE REFERENCIA ***
     * "construtora que vende apartamentos"     -> "Fisico -> Industria / fabricante"
     * "incorporadora imobiliaria"              -> "Fisico -> Industria / fabricante"
     * "industria de moveis sob medida"         -> "Fisico -> Industria / fabricante"
     * "loja online de cosmeticos proprios"     -> "Fisico -> Industria / fabricante" (se fabrica) OU "Fisico -> E-commerce / varejo online" (se revende)
     * "marketplace de imoveis (sem estoque)"   -> "Digital -> Plataforma / Marketplace"
     * "SaaS de automacao WhatsApp"             -> "Digital -> SaaS / assinatura"
     * "escritorio de advocacia"                -> "Servico -> Consultoria / assessoria"
     * "agencia de marketing digital"           -> "Servico -> Agencia"
     * "clinica medica/odontologica"            -> "Servico -> Consultoria / assessoria" (e o melhor encaixe quando o entregavel e atendimento humano)

   Escolha a opcao que melhor descreve o NEGOCIO PRINCIPAL do site (ignore receitas secundarias).
- "audiencia_alvos": APENAS uma lista CURTA (entre 6 e 10 itens) separada por virgulas, contendo
   (a) cargos de tomadores de decisao,
   (b) tipos de empresas (ex: PMEs, agencias, escritorios) e
   (c) nichos/segmentos especificos que mais usam o produto/servico.
   Cada item com no maximo 4 palavras. SEM descricoes, SEM explicacoes, SEM introducoes,
   SEM repeticao de conceitos (nao incluir "clinicas medicas" E "clinicas de saude" juntos).
   Exemplo de saida ideal: "PMEs, gestores comerciais, advogados, clinicas de saude, profissionais de wellness, instituicoes de educacao, equipes de suporte e CS".
- Adicione tambem "termos_raiz": lista de 3 a 6 substantivos genericos que descrevem o produto/servico do site (ex: ["construtora", "apartamento", "empreendimento", "imovel"]). Esses substantivos devem aparecer no TITLE/META/H1/NAV.

Responda SOMENTE o JSON, sem texto antes ou depois.
"""
    raw = call_openrouter(prompt, max_tokens=1200, temperature=0.2)
    data = extract_json(raw)

    for k, _ in BRIEFING_FIELDS:
        data.setdefault(k, "")
    data.setdefault("cidades", [])
    data.setdefault("termos_raiz", [])
    data["escopo"] = "Local" if str(data.get("escopo", "")).strip().lower().startswith("local") else "Nacional"
    if isinstance(data.get("cidades"), str):
        data["cidades"] = [c.strip() for c in re.split(r"[,;]", data["cidades"]) if c.strip()]
    if isinstance(data.get("termos_raiz"), str):
        data["termos_raiz"] = [t.strip() for t in re.split(r"[,;]", data["termos_raiz"]) if t.strip()]
    return data


# ---------------------------------------------------------------------------
# ETAPA 1C: LLM PERSONAS + POTENCIAL DE CANAIS
# ---------------------------------------------------------------------------

def llm_personas_canais(url: str, site: dict, briefing: dict) -> dict:
    """Gera 2-3 personas + matriz de potencial de canais de marketing.

    Retorno (dict):
      {
        "personas": [
          {"nome": str, "idade_perfil": str, "descricao": str, "onde_busca": str},
          ...
        ],
        "canais": [
          {"canal": str, "conversao": "Alta|Media|Baixa",
           "velocidade": "Alta|Media|Baixa", "custo": "Alto|Medio|Baixo",
           "justificativa": str},
          ...
        ]
      }
    """
    print("[1C] Gerando personas e potencial de canais com LLM ...")
    contexto = (
        f"Nicho: {briefing.get('nicho', '')}\n"
        f"Nicho principal: {briefing.get('nicho_principal', '')}\n"
        f"Publico alvo: {briefing.get('publico_alvo', '')}\n"
        f"B2B/B2C: {briefing.get('b2b_b2c', '')}\n"
        f"Escopo: {briefing.get('escopo', '')}\n"
        f"Tipo: {briefing.get('tipo_negocio', '')}\n"
        f"Modelo de Produto: {briefing.get('modelo_produto', '')}\n"
        f"Dores: {briefing.get('dores', '')}\n"
    )

    prompt = f"""Voce e um estrategista de marketing digital. A partir do briefing abaixo da empresa, gere:
1) De 2 a 3 PERSONAS detalhadas (perfis de comprador/decisor).
2) Uma MATRIZ DE POTENCIAL DE CANAIS DE MARKETING com 6 a 8 canais relevantes para esse negocio,
   classificando cada canal em Conversao, Velocidade de Retorno e Custo/Esforco com Alta/Media/Baixa
   (ou Alto/Medio/Baixo para custo) e uma justificativa curta focada nas personas.

URL: {url}

BRIEFING:
{contexto}

CONTEUDO DO SITE (resumo):
{site_signals_text(site)[:4000]}

RETORNE APENAS JSON VALIDO no formato exato abaixo (sem texto antes ou depois):
{{
  "personas": [
    {{
      "nome": "Persona 1: <nome curto do papel/cargo>",
      "idade_perfil": "Faixa etaria e descricao do perfil em 1 frase (ex: '30 a 50 anos. Perfil pragmatico, focado em crescimento.')",
      "descricao": "Frase complementar com motivacoes/objetivos principais",
      "onde_busca": "Onde essa persona busca informacao (canais, redes, sites, comunidades). Frase curta."
    }}
  ],
  "canais": [
    {{
      "canal": "Nome do canal (ex: 'SEO / Inbound Marketing', 'Trafego Pago (Meta/Google)', 'YouTube', 'LinkedIn Ads / Organic', 'TikTok', 'Indicacao / Networking', 'E-mail Marketing', 'Eventos / Feiras')",
      "conversao": "Alta|Media|Baixa",
      "velocidade": "Alta|Media|Baixa",
      "custo": "Alto|Medio|Baixo",
      "justificativa": "Justificativa estrategica curta (1-2 frases) referenciando a(s) persona(s)"
    }}
  ]
}}

REGRAS:
- Use canais REALMENTE relevantes para o negocio (nao force todos os 8 se nao fizer sentido).
- Para B2B priorize: LinkedIn, SEO, Indicacao, Eventos, Tragefo Pago, YouTube, E-mail.
- Para B2C local (servicos): Google Meu Negocio, Indicacao, SEO Local, Instagram, Trafego Pago, WhatsApp.
- Para B2C nacional/ecommerce: SEO, Trafego Pago, Influenciadores, Instagram, TikTok, YouTube, E-mail.
- "conversao" = potencial real de virar venda.
- "velocidade" = quao rapido aparecem resultados.
- "custo" = investimento financeiro + esforco operacional combinados.
- "justificativa": 1-2 frases citando a persona principal beneficiada.
"""
    raw = call_openrouter(prompt, max_tokens=2200, temperature=0.3)
    try:
        data = extract_json(raw)
    except Exception as e:
        print(f"[1C] Falha ao parsear JSON: {e}. Usando estrutura vazia.")
        return {"personas": [], "canais": []}

    personas = data.get("personas") or []
    canais = data.get("canais") or []

    def _norm(v, allowed):
        s = str(v or "").strip()
        for a in allowed:
            if s.lower().startswith(a.lower()[:3]):
                return a
        return s or "—"

    for p in personas:
        for k in ("nome", "idade_perfil", "descricao", "onde_busca"):
            p.setdefault(k, "")
    for c in canais:
        c.setdefault("canal", "")
        c["conversao"] = _norm(c.get("conversao"), ["Alta", "Media", "Baixa"])
        c["velocidade"] = _norm(c.get("velocidade"), ["Alta", "Media", "Baixa"])
        c["custo"] = _norm(c.get("custo"), ["Alto", "Medio", "Baixo"])
        c.setdefault("justificativa", "")

    print(f"[1C] {len(personas)} persona(s) | {len(canais)} canal(is) gerado(s)")
    return {"personas": personas, "canais": canais}


# ---------------------------------------------------------------------------
# ETAPA 2: LLM GERA KEYWORDS
# ---------------------------------------------------------------------------

# Gatilhos informacionais (TOPO de funil): aprendizado/duvida, sem intencao de compra
_TOPO_PREFIXES = (
    "como ", "o que ", "o q ", "qual ", "quais ", "por que ", "porque ",
    "quando ", "onde ", "quem ", "para que ", "para quem ",
)
_TOPO_CONTAINS = (
    "vale a pena", "quanto custa", "quanto e", "quanto fica",
    "diferenca entre", "diferenca de", "diferenca do",
    "tipos de", "tipo de", "significado", "significa",
    "como funciona", "como escolher", "como contratar", "como cancelar",
    "passo a passo", "guia ", " guia", "tutorial",
    "vantagens", "desvantagens", "beneficios", "prós e contras", "pros e contras",
    "exemplo", "ideias", "dicas", "tudo sobre", "o que sao", "o que são",
    "como saber", "como identificar", "como usar", "como fazer",
    "vale apena", "compensa", "preciso ", " precisa",
)

# Gatilhos comerciais explicitos (FUNDO de funil)
_FUNDO_CONTAINS = (
    "comprar", "preco", "preço", "preços", "precos", "valor", "valores",
    "melhor ", "melhores ", "top ", "ranking",
    "barato", "barata", "baratos", "baratas",
    "contratar", "contratacao", "contratação", "assinar", "assinatura",
    "cotacao", "cotação", "cotar", "orcamento", "orçamento",
    "venda", "a venda", "à venda", "ofertas", "oferta",
    "promocao", "promoção", "desconto", "cupom",
    "loja", "lojas", "perto de mim", "agora",
    "plano ", "planos ", " plano", " planos",
    "marca ", "marcas",
)


# Stopwords/conectores PT-BR que NAO mudam a intencao da keyword
_KW_STOPWORDS = {
    "de", "do", "da", "dos", "das", "em", "no", "na", "nos", "nas",
    "para", "pra", "por", "pelo", "pela", "pelos", "pelas",
    "com", "sem", "e", "ou", "o", "a", "os", "as",
    "um", "uma", "uns", "umas", "ao", "aos",
}


def _stem_token(t: str) -> str:
    """Stemmer minimalista PT-BR: remove apenas o sufixo de plural mais comum
    (s/es) para que 'plano'/'planos' e 'gato'/'gatos' sejam tratados como o
    mesmo token. Conservador: nao mexe em tokens curtos."""
    if len(t) >= 5 and t.endswith("es") and t[-3] in "rsz":
        return t[:-2]  # "mares" -> "mar", "luzes" -> "luz", "paises" -> "pais"
    if len(t) >= 4 and t.endswith("s") and not t.endswith("us"):
        return t[:-1]  # "planos" -> "plano", "gatos" -> "gato"
    return t


def keyword_signature(kw: str) -> frozenset:
    """Assinatura de intencao de uma keyword. Duas keywords com a MESMA
    assinatura sao tratadas como variacoes da mesma intencao.

    Algoritmo:
    1. lowercase + remove acentos (ja vem assim do pipeline)
    2. tokeniza, ignora conectores ('de', 'para', 'em', etc.)
    3. aplica stemming simples (remove plural 's'/'es')
    4. retorna frozenset (ordem nao importa)

    Exemplos:
        'plano de saude pet'         -> {plano, saude, pet}
        'planos de saude pet'        -> {plano, saude, pet}    (mesma)
        'planos de saude para pet'   -> {plano, saude, pet}    (mesma)
        'planos de saude para pets'  -> {plano, saude, pet}    (mesma)
        'plano de saude para gato'   -> {plano, saude, gato}   (diferente)
    """
    if not kw:
        return frozenset()
    s = unicodedata.normalize("NFKD", kw.lower()).encode("ascii", "ignore").decode("ascii")
    tokens = re.findall(r"[a-z0-9]+", s)
    stems = {_stem_token(t) for t in tokens if t and t not in _KW_STOPWORDS}
    return frozenset(stems)


def dedupe_keywords_by_intent(sorted_kws: list, n: int) -> tuple:
    """Pega ate N keywords distintas semanticamente de uma lista JA ORDENADA por
    prioridade (volume desc). Variacoes morfologicas sao descartadas (mantem
    apenas a primeira de cada grupo).

    Retorna (selected, dropped) onde dropped lista as variacoes descartadas
    com o motivo (qual keyword venceu)."""
    selected = []
    selected_sigs = []
    dropped = []
    for k in sorted_kws:
        sig = keyword_signature(k["keyword"])
        match_idx = None
        for i, s in enumerate(selected_sigs):
            if sig == s:
                match_idx = i
                break
        if match_idx is not None:
            dropped.append({"keyword": k["keyword"], "duplicate_of": selected[match_idx]["keyword"]})
            continue
        selected.append(k)
        selected_sigs.append(sig)
        if len(selected) >= n:
            break
    return selected, dropped


def classify_funnel(keyword: str, escopo: str, termos_raiz: list) -> str:
    """Re-classifica deterministicamente uma keyword como 'local', 'topo' ou 'fundo'.
    Aplica sobre a saida do LLM para garantir consistencia (ex: 'plano de saude pet' e
    'planos de saude pet' devem cair no MESMO funil)."""
    if escopo == "Local":
        return "local"

    kw = (keyword or "").lower().strip()
    if not kw:
        return "topo"

    # 1) Prefixos informacionais sao TOPO mesmo que mencionem o produto
    for p in _TOPO_PREFIXES:
        if kw.startswith(p):
            return "topo"
    # 2) Padroes informacionais em qualquer posicao
    for c in _TOPO_CONTAINS:
        if c in kw:
            return "topo"
    # 3) Sinais comerciais explicitos
    for c in _FUNDO_CONTAINS:
        if c in kw:
            return "fundo"
    # 4) Default = FUNDO. Usuario buscando o produto/servico diretamente
    #    (ex: "plano de saude pet", "seguro animal") tem intencao comercial.
    #    Topo de funil precisa de gatilho informacional explicito.
    return "fundo"


def llm_generate_keywords(briefing: dict, site: dict) -> list:
    escopo = briefing["escopo"]
    cidades = briefing.get("cidades") or []
    termos_raiz = briefing.get("termos_raiz") or []
    site_excerpt = (
        f"TITLE: {site.get('title','')}\n"
        f"META: {site.get('meta_description','')}\n"
        f"H1: {' | '.join(site.get('h1',[]))}\n"
        f"NAV: {' | '.join(site.get('nav',[]))}"
    )

    final_cap = MAX_LOCAL if escopo == "Local" else MAX_NACIONAL
    overgen = int(final_cap * OVERGEN_FACTOR)

    if escopo == "Local":
        cidades_str = ", ".join(cidades) if cidades else "(cidades a definir)"
        raiz_str = ", ".join(termos_raiz) if termos_raiz else "(extrair do TITLE/H1/NAV abaixo)"
        regra = f"""Gere EXATAMENTE {overgen} palavras-chave de SEO LOCAL combinando SERVICO/PRODUTO + CIDADE/REGIAO.

TERMOS-RAIZ obrigatorios (priorize estes substantivos genericos): {raiz_str}
CIDADES alvo: {cidades_str}

REGRAS DE OURO (siga sem excecao):
1. Comece pelos termos MAIS GENERICOS e de MAIOR VOLUME. Para cada termo-raiz, gere combinacoes nesta ordem:
   - "<raiz> <cidade>"            (ex: "construtora itajai")
   - "<raiz> em <cidade>"          (ex: "construtoras em itajai")
   - "<raiz>s <cidade>"            (plural, ex: "construtoras itajai")
   - "<raiz> <cidade> sc"          (com estado)
   - "<raiz> <bairro>"             (so se o bairro aparece no site, ex: "apartamento praia brava")
2. NAO adicione adjetivos qualificadores como "luxo", "alto padrao", "premium", "exclusivo", "alto-padrao", "de luxo" EXCETO se essa expressao aparecer LITERAL no TITLE/H1/NAV do site abaixo.
3. Use lowercase, sem acento, sem pontuacao.
4. Inclua tambem 2-3 variacoes de intencao comercial direta: "comprar <raiz> <cidade>", "<raiz> a venda <cidade>".
5. Evite cauda longa muito especifica (ex: nao gere "apartamento com 3 quartos vista mar em itajai sc 2026").

SINAIS DIRETOS DO SITE (use como evidencia para os termos):
{site_excerpt}

Retorne JSON com lista "keywords" de objetos:
[{{"keyword": "...", "funil": "local", "categoria": "<termo-raiz>"}}]"""
    else:
        raiz_str = ", ".join(termos_raiz) if termos_raiz else "(extrair do TITLE/H1/NAV abaixo)"
        regra = f"""Gere EXATAMENTE {overgen} palavras-chave de SEO NACIONAL divididas em:
- FUNDO de funil (~60%): usuario com intencao comercial direta.
  * O TERMO RAIZ PURO ja eh fundo de funil (ex: "plano de saude pet", "construtora itajai", "seguro pet").
  * Variacoes plurais/morfologicas pertencem AO MESMO FUNIL do termo raiz (ex: "planos de saude pet" = fundo, igual a "plano de saude pet").
  * Modificadores comerciais: "comprar", "preco", "valor", "melhor", "barato", "contratar", "cotacao", "perto de mim".
- TOPO de funil (~40%): usuario buscando informacao/aprendendo sobre o produto.
  * Inicia com pronome interrogativo: "como", "o que e", "qual", "quanto custa", "quando", "por que".
  * Ou contem: "vale a pena", "diferenca entre", "tipos de", "como funciona", "como escolher", "vantagens", "beneficios", "guia", "passo a passo".

REGRA DE CONSISTENCIA: variacoes da mesma intencao DEVEM ter o mesmo funil. Se voce classificar "plano de saude pet" como fundo, "planos de saude pet" tambem eh fundo.

TERMOS-RAIZ obrigatorios: {raiz_str}
Nicho: "{briefing.get('nicho_principal','')}" | Secundario: "{briefing.get('nicho_secundario','')}"
Dores: {briefing.get('dores','')}

REGRAS DE OURO:
1. Comece pelos termos MAIS GENERICOS e de MAIOR VOLUME, depois variacoes.
2. NAO adicione adjetivos qualificadores ("luxo", "premium", "exclusivo", etc.) EXCETO se aparecerem literais no TITLE/H1/NAV abaixo.
3. Use lowercase, sem acento.

SINAIS DIRETOS DO SITE:
{site_excerpt}

Retorne JSON com lista "keywords" de objetos:
[{{"keyword": "...", "funil": "topo" ou "fundo", "categoria": "<termo-raiz>"}}]"""

    prompt = f"""Voce e especialista em pesquisa de palavras-chave SEO no Brasil. Voce conhece como termos sao realmente pesquisados no Google e prioriza volume real ao inves de termos artificiais.

TAREFA:
{regra}

Responda APENAS JSON valido no formato: {{"keywords": [...]}}"""

    print(f"[2] Gerando {overgen} keywords ({escopo}, cap final {final_cap}) ...")
    raw = call_openrouter(prompt, max_tokens=3500, temperature=0.2)
    data = extract_json(raw)
    kws = data.get("keywords", [])

    seen, out = set(), []
    for item in kws:
        kw = str(item.get("keyword", "")).strip().lower()
        kw = unicodedata.normalize("NFKD", kw).encode("ascii", "ignore").decode("ascii")
        kw = re.sub(r"\s+", " ", kw)
        if not kw or kw in seen:
            continue
        seen.add(kw)
        # Re-classificacao deterministica do funil (sobrescreve o LLM para
        # garantir consistencia entre variacoes morfologicas da mesma intencao).
        funil = classify_funnel(kw, escopo, termos_raiz)
        out.append({
            "keyword": kw,
            "funil": funil,
            "categoria": str(item.get("categoria", "")).strip(),
        })
    print(f"[2] {len(out)} keywords geradas (antes do filtro de volume)")
    return out


def filter_keywords_by_volume(keywords: list, volumes: dict, escopo: str,
                                brand: str = "", generic_terms: list = None) -> list:
    """Cap final: somente keywords com volume > 0 (ordenadas por volume desc).

    Filtros aplicados:
      - drop keywords com volume = 0 ou ausente (sem busca)
      - drop keywords que casam com a marca do cliente (nome da empresa)
    Limite final: 20 (local) ou 50 (nacional).
    """
    cap = MAX_LOCAL if escopo == "Local" else MAX_NACIONAL
    brand = (brand or "").strip().lower()

    n_drop_brand = 0
    n_drop_zero = 0
    with_vol = []
    for k in keywords:
        v = (volumes.get(k["keyword"], {}) or {}).get("volume_medio") or 0
        if v <= 0:
            n_drop_zero += 1
            continue
        kw_norm = (k.get("keyword") or "").lower()
        if brand and is_brand_keyword(kw_norm, brand, generic_terms=generic_terms):
            n_drop_brand += 1
            continue
        with_vol.append((v, k))
    with_vol.sort(key=lambda x: x[0], reverse=True)
    out = [k for _, k in with_vol][:cap]
    print(f"[2b] Filtrado para {len(out)} keywords (todas com volume > 0)"
          f" | descartadas: {n_drop_zero} sem volume, {n_drop_brand} de marca")
    return out


# ---------------------------------------------------------------------------
# ETAPA 3: DATAFORSEO
# ---------------------------------------------------------------------------

def fetch_search_volume(keywords: list) -> dict:
    if not keywords:
        return {}
    print(f"[3] Buscando volume no DataForSEO ({len(keywords)} keywords) ...")
    auth = requests.auth.HTTPBasicAuth(DATAFORSEO_USER, DATAFORSEO_PASS)
    today = datetime.today()
    date_to = today.strftime("%Y-%m-%d")
    date_from = today.replace(year=today.year - 3).strftime("%Y-%m-%d")

    payload = [{
        "location_code": LOCATION_BRAZIL,
        "language_code": "pt",
        "keywords": keywords,
        "date_from": date_from,
        "date_to": date_to,
        "search_partners": False,
    }]
    r = requests.post(DATAFORSEO_VOLUME_URL, auth=auth, json=payload, timeout=120)
    r.raise_for_status()
    js = r.json()
    if js.get("status_code") != 20000:
        print(f"[3] API erro: {js.get('status_code')} {js.get('status_message')}")
        return {}

    results = (js["tasks"][0].get("result") or [])
    by_kw = {}
    for item in results:
        kw = (item.get("keyword") or "").lower()
        vol_avg = item.get("search_volume")
        cpc = item.get("cpc")
        competition = item.get("competition")

        recent, year_ago = None, None
        monthly = item.get("monthly_searches") or []
        if monthly:
            sm = sorted(monthly, key=lambda m: (m["year"], m["month"]))
            recent_entry = sm[-1]
            recent = recent_entry["search_volume"]
            ry, rm = recent_entry["year"] - 1, recent_entry["month"]
            for x in sm:
                if x["year"] == ry and x["month"] == rm:
                    year_ago = x["search_volume"]
                    break

        growth = None
        if recent is not None and year_ago not in (None, 0):
            growth = round(((recent - year_ago) / year_ago) * 100, 2)

        by_kw[kw] = {
            "volume_medio": vol_avg,
            "volume_recente": recent,
            "volume_ano_anterior": year_ago,
            "crescimento_1a_pct": growth,
            "cpc": cpc,
            "competition": competition,
        }
    print(f"[3] {len(by_kw)} volumes obtidos")
    return by_kw


# ---------------------------------------------------------------------------
# ETAPA 3B: DATAFORSEO SERP (posicao + top concorrentes)
# ---------------------------------------------------------------------------

def normalize_domain(d: str) -> str:
    d = (d or "").lower().strip()
    if d.startswith("http"):
        d = urlparse(d).netloc
    if d.startswith("www."):
        d = d[4:]
    return d


def _fetch_one_serp(kw: str, depth: int) -> tuple:
    """A API DataForSEO SERP advanced/live aceita apenas 1 task por request."""
    auth = requests.auth.HTTPBasicAuth(DATAFORSEO_USER, DATAFORSEO_PASS)
    payload = [{
        "keyword": kw,
        "location_code": LOCATION_BRAZIL,
        "language_code": "pt",
        "device": "desktop",
        "os": "windows",
        "depth": depth,
    }]
    try:
        r = requests.post(DATAFORSEO_SERP_URL, auth=auth, json=payload, timeout=120)
        r.raise_for_status()
        js = r.json()
    except Exception as e:
        return kw, [], f"erro request: {e}"

    items_out = []
    for task in js.get("tasks") or []:
        if task.get("status_code") != 20000:
            return kw, [], task.get("status_message") or "task error"
        for r0 in task.get("result") or []:
            for it in r0.get("items") or []:
                if it.get("type") != "organic":
                    continue
                items_out.append({
                    "rank": it.get("rank_absolute"),
                    "domain": normalize_domain(it.get("domain") or ""),
                    "url": it.get("url") or "",
                    "title": it.get("title") or "",
                    "description": it.get("description") or "",
                })
    items_out.sort(key=lambda x: x["rank"] or 999)
    return kw, items_out, None


def fetch_serps(keywords: list, depth: int = SERP_DEPTH, workers: int = 6) -> dict:
    if not keywords:
        return {}
    print(f"[3B] Buscando SERPs no DataForSEO ({len(keywords)} keywords, {workers} threads) ...")
    by_kw = {}
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_fetch_one_serp, kw, depth): kw for kw in keywords}
        for fut in as_completed(futures):
            kw, items, err = fut.result()
            if err:
                print(f"[3B]   ! {kw!r}: {err}")
            by_kw[kw] = items
    print(f"[3B] {sum(1 for v in by_kw.values() if v)}/{len(keywords)} SERPs com resultados organicos")
    return by_kw


def find_position(items: list, client_domain: str):
    cd = normalize_domain(client_domain)
    for it in items:
        if it["domain"] == cd:
            return it["rank"]
    return None


# ---------------------------------------------------------------------------
# ETAPA 4: SIMILARIDADE DO TOP 3 NAS TOP 5 KEYWORDS
# ---------------------------------------------------------------------------

SIMILARITY_LEVELS = {"alto": "Alto", "alta": "Alto", "high": "Alto",
                     "medio": "Medio", "media": "Medio", "medium": "Medio", "médio": "Medio", "média": "Medio",
                     "baixo": "Baixo", "baixa": "Baixo", "low": "Baixo"}


def scrape_url_text(url: str, max_chars: int = 4000) -> str:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    }
    try:
        r = requests.get(url, headers=headers, timeout=20, allow_redirects=True)
        r.raise_for_status()
    except Exception:
        return ""
    soup = BeautifulSoup(r.text, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    title = (soup.title.string.strip() if soup.title and soup.title.string else "")
    m = soup.find("meta", attrs={"name": "description"})
    desc = m["content"].strip() if m and m.get("content") else ""
    h1s = " | ".join(h.get_text(strip=True) for h in soup.find_all("h1")[:5])
    body = re.sub(r"\s+", " ", soup.get_text(separator=" ", strip=True))
    return f"TITLE: {title}\nMETA: {desc}\nH1: {h1s}\nBODY: {body}"[:max_chars]


PROFILE_LEVELS = {"especialista": "Especialista", "generalista": "Generalista"}


def _parse_profile(val: str) -> str:
    """Normaliza 'Especialista'/'Generalista' a partir de string suja do LLM."""
    if not val:
        return ""
    v = unicodedata.normalize("NFKD", val.lower()).encode("ascii", "ignore").decode("ascii")
    v = re.sub(r"[^a-z]", "", v)
    return PROFILE_LEVELS.get(v, val.capitalize() if val else "")


def analyze_similarity(comp_url: str, client_url: str, briefing: dict) -> dict:
    content = scrape_url_text(comp_url)
    if not content:
        return {"nicho": "—", "similaridade": "—", "perfil": "—", "motivo": "Falha ao acessar URL"}

    client_niche = (
        f"Nicho: {briefing.get('nicho','')} | "
        f"Nicho principal: {briefing.get('nicho_principal','')} | "
        f"Tipo: {briefing.get('tipo_negocio','')} | "
        f"Publico: {briefing.get('publico_alvo','')} | "
        f"Escopo: {briefing.get('escopo','')}"
    )

    prompt = f"""Analise o conteudo do website abaixo e compare com o cliente principal.

WEBSITE A ANALISAR:
URL: {comp_url}
Conteudo: {content}

CLIENTE PRINCIPAL:
URL: {client_url}
Nicho/Descricao: {client_niche}

Responda EXATAMENTE neste formato (4 linhas):

NICHO: [o nicho principal deste website em poucas palavras]
SIMILARIDADE: [Baixo, Medio ou Alto]
PERFIL: [Especialista ou Generalista]
MOTIVO: [explicacao breve, 1 frase, considerando nicho, publico-alvo, servicos E perfil]

Criterios de SIMILARIDADE:
- Alto: Mesmo nicho, mesmo publico-alvo, servicos muito similares (concorrente direto que disputa os mesmos clientes)
- Medio: Nicho relacionado, alguma sobreposicao de publico ou servicos (ex: portal de imoveis vs construtora especifica)
- Baixo: Nicho diferente, pouca ou nenhuma relacao (ex: site de noticias, agregador generico, blog)

Criterios de PERFIL (em relacao ao nicho do cliente "{briefing.get('nicho_principal','')}"):
- Especialista: o website tem foco UNICO ou PRINCIPAL no mesmo nicho do cliente. Ex: cliente vende "plano de saude pet" e o concorrente TAMBEM so vende plano de saude pet (sem outros produtos relevantes).
- Generalista: o website oferece MULTIPLOS produtos/servicos e o nicho do cliente eh apenas UM entre varios (e nao necessariamente o principal). Ex: cliente vende plano pet e o concorrente eh uma seguradora gigante com seguro auto, vida, residencial, cartao, etc., onde plano pet eh so um produto secundario.

Use o conteudo do BODY/H1/MENU para inferir o perfil. Se o site tem menu com varias categorias nao-relacionadas ao nicho do cliente, eh Generalista. Se todo o conteudo gira em torno do mesmo nicho, eh Especialista.

Responda em portugues do Brasil."""

    try:
        raw = call_openrouter(prompt, max_tokens=350, temperature=0.2)
    except Exception as e:
        return {"nicho": "—", "similaridade": "—", "perfil": "—", "motivo": f"Erro LLM: {e}"}

    out = {"nicho": "", "similaridade": "", "perfil": "", "motivo": ""}
    for line in raw.splitlines():
        line = line.strip().lstrip("*-•").strip()
        low = line.lower()
        if low.startswith("nicho:"):
            out["nicho"] = line.split(":", 1)[1].strip()
        elif low.startswith("similaridade:"):
            val = line.split(":", 1)[1].strip().lower()
            val = unicodedata.normalize("NFKD", val).encode("ascii", "ignore").decode("ascii")
            val = re.sub(r"[^a-z]", "", val)
            out["similaridade"] = SIMILARITY_LEVELS.get(val, val.capitalize() or "—")
        elif low.startswith("perfil:"):
            out["perfil"] = _parse_profile(line.split(":", 1)[1].strip())
        elif low.startswith("motivo:"):
            out["motivo"] = line.split(":", 1)[1].strip()
    if not out["similaridade"]:
        out["similaridade"] = "—"
    if not out["perfil"]:
        out["perfil"] = "—"
    return out


def classify_client_profile(url: str, site: dict, briefing: dict) -> str:
    """Classifica o proprio cliente como Especialista ou Generalista usando
    o site ja scrapeado no briefing (evita re-scraping)."""
    site_text = (
        f"TITLE: {site.get('title','')}\n"
        f"META: {site.get('meta_description','')}\n"
        f"H1: {' | '.join(site.get('h1',[]))}\n"
        f"H2: {' | '.join(site.get('h2',[]))}\n"
        f"NAV: {' | '.join(site.get('nav',[]))}\n"
        f"BODY: {(site.get('body','') or '')[:1500]}"
    )
    nicho_principal = briefing.get("nicho_principal", "")
    prompt = f"""Classifique o website abaixo como Especialista ou Generalista no nicho "{nicho_principal}".

WEBSITE:
URL: {url}
Conteudo:
{site_text}

Criterios:
- Especialista: foco UNICO ou PRINCIPAL no nicho "{nicho_principal}". Quase todo o conteudo, menu e produtos giram em torno desse nicho.
- Generalista: oferece MULTIPLOS produtos/servicos e o nicho "{nicho_principal}" eh apenas UM entre varios (ex: uma seguradora gigante que tambem vende esse produto).

Responda EXATAMENTE 1 linha:
PERFIL: [Especialista ou Generalista]"""
    try:
        raw = call_openrouter(prompt, max_tokens=60, temperature=0.1)
    except Exception:
        return "—"
    for line in raw.splitlines():
        low = line.strip().lower()
        if low.startswith("perfil:"):
            return _parse_profile(line.split(":", 1)[1].strip())
    return _parse_profile(raw.strip())


# ---------------------------------------------------------------------------
# ETAPA 6: BACKUP SEMRUSH (top 30 keywords de um concorrente Alto)
# ---------------------------------------------------------------------------

def _semrush_request_id() -> str:
    k = "".join(random.choices(ascii_lowercase + digits, k=33))
    return k[:8] + "-" + k[8:12] + "-" + k[12:16] + "-" + k[16:]


def fetch_semrush_top_keywords(domain: str, limit: int = SEMRUSH_TOP_N,
                                db: str = SEMRUSH_DB) -> list:
    """Top N keywords organicas de um dominio via Semrush JSON-RPC (organic.Positions).
    Usa autenticacao por apiKey (sem cookies)."""
    clean = normalize_domain(domain)
    print(f"[6] Buscando top {limit} keywords de {clean} no Semrush (db={db}) ...")
    payload = {
        "id": 1, "jsonrpc": "2.0", "method": "organic.Positions",
        "params": {
            "request_id": _semrush_request_id(),
            "report": "organic.positions",
            "args": {
                "database": db, "searchItem": clean, "searchType": "domain",
                "filter": {}, "dateType": "daily", "positionsType": "all",
                "display": {
                    "order": {"field": "trafficPercent", "direction": "desc"},
                    "page": 1, "pageSize": limit,
                },
            },
            "apiKey": SEMRUSH_API_KEY,
        },
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Content-Type": "application/json",
        "Origin": "https://www.semrush.com",
        "Referer": "https://www.semrush.com/analytics/organic/positions/",
    }
    try:
        r = requests.post(SEMRUSH_RPC_URL, json=payload, headers=headers, timeout=45)
        r.raise_for_status()
        js = r.json()
    except Exception as e:
        print(f"[6] Erro Semrush: {e}")
        return []
    if isinstance(js, dict) and "error" in js:
        print(f"[6] Erro RPC: {js['error']}")
        return []
    result = js.get("result") if isinstance(js, dict) else None
    if not isinstance(result, list):
        print(f"[6] Resultado inesperado: {str(js)[:200]}")
        return []
    out = []
    for item in result:
        if not isinstance(item, dict):
            continue
        out.append({
            "phrase": item.get("phrase", ""),
            "position": item.get("position"),
            "volume": item.get("volume"),
            "kd": item.get("keywordDifficulty"),
            "traffic": item.get("traffic"),
            "cpc": item.get("cpc"),
            "url": item.get("url", ""),
        })
    print(f"[6] {len(out)} keywords obtidas")
    return out


def fetch_semrush_monthly_trend(domain: str, db: str = SEMRUSH_DB) -> list:
    """Serie historica mensal de trafego organico via Semrush JSON-RPC
    (organic.MonthlyTrend, report=organic.overview).

    Retorna lista de dicts com:
        - 'ts'         : unix timestamp do mes
        - 'date'       : 'YYYY-MM' formatado
        - 'branded'    : organicTrafficBranded
        - 'non_branded': organicTrafficNonBranded
        - 'total'      : organicTraffic
        - 'positions'  : organicPositions
    """
    clean = normalize_domain(domain)
    payload = {
        "id": 1, "jsonrpc": "2.0", "method": "organic.MonthlyTrend",
        "params": {
            "request_id": _semrush_request_id(),
            "report": "organic.overview",
            "args": {
                "database": db, "searchItem": clean,
                "searchType": "domain", "filter": {},
            },
            "apiKey": SEMRUSH_API_KEY,
        },
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Content-Type": "application/json",
        "Origin": "https://www.semrush.com",
        "Referer": "https://www.semrush.com/analytics/overview/",
    }
    try:
        r = requests.post(SEMRUSH_RPC_URL, json=payload, headers=headers, timeout=45)
        r.raise_for_status()
        js = r.json()
    except Exception as e:
        print(f"[7] Erro Semrush MonthlyTrend ({clean}): {e}")
        return []
    if isinstance(js, dict) and "error" in js:
        print(f"[7] Erro RPC MonthlyTrend ({clean}): {js['error']}")
        return []
    result = js.get("result") if isinstance(js, dict) else None
    if not isinstance(result, list):
        return []
    out = []
    for it in result:
        if not isinstance(it, dict):
            continue
        ts = it.get("date")
        if not ts:
            continue
        dt = datetime.fromtimestamp(int(ts))
        out.append({
            "ts": int(ts),
            "date": dt.strftime("%Y-%m"),
            "branded": int(it.get("organicTrafficBranded") or 0),
            "non_branded": int(it.get("organicTrafficNonBranded") or 0),
            "total": int(it.get("organicTraffic") or 0),
            "positions": int(it.get("organicPositions") or 0),
        })
    out.sort(key=lambda x: x["ts"])
    return out


def _pick_trend_points(series: list) -> dict:
    """A partir de uma serie mensal, pega o ponto mais recente ('current')
    e o ponto ~12 meses antes ('year_ago')."""
    if not series:
        return {"current": None, "year_ago": None}
    current = series[-1]
    target_ts = current["ts"] - 365 * 86400
    # acha o ponto mais proximo do alvo (12 meses antes)
    year_ago = min(series[:-1], key=lambda p: abs(p["ts"] - target_ts)) if len(series) > 1 else None
    return {"current": current, "year_ago": year_ago}


def build_traffic_comparison(client_domain: str, similarity_rows: list,
                              max_competitors: int = 3,
                              client_profile: str = "—") -> dict:
    """Etapa 7: pega ate N concorrentes Alto distintos da analise de similaridade
    e busca Semrush MonthlyTrend para CLIENTE + concorrentes.

    Retorna dict:
        {
          'rows': [
            {'domain': ..., 'role': 'cliente'|'concorrente',
             'current': {...}, 'year_ago': {...},
             'growth_total_pct': ..., 'growth_nonbranded_pct': ...},
            ...
          ],
          'current_date': 'YYYY-MM',
          'year_ago_date': 'YYYY-MM',
        }"""
    client = normalize_domain(client_domain)

    # concorrentes Alto unicos preservando ordem de aparicao (top SERPs primeiro)
    # e mapa de perfil por dominio (vindo da analise de similaridade)
    seen_doms = set()
    alto_doms = []
    profile_by_domain = {client: client_profile}
    for s in similarity_rows:
        d = normalize_domain(s.get("domain", ""))
        if d and d not in profile_by_domain and s.get("perfil"):
            profile_by_domain[d] = s.get("perfil")
        if s.get("similaridade") != "Alto":
            continue
        if not d or d == client or d in seen_doms:
            continue
        seen_doms.add(d)
        alto_doms.append(d)
        if len(alto_doms) >= max_competitors:
            break

    targets = [(client, "cliente")] + [(d, "concorrente") for d in alto_doms]
    print(f"[7] Comparativo de trafego SEO: cliente + {len(alto_doms)} concorrente(s) Alto")
    for d, role in targets:
        print(f"      - {role}: {d} ({profile_by_domain.get(d, '—')})")

    rows = []
    cur_dates, ya_dates = set(), set()
    for dom, role in targets:
        series = fetch_semrush_monthly_trend(dom)
        if not series:
            print(f"[7]   {dom}: sem dados")
            rows.append({"domain": dom, "role": role, "current": None, "year_ago": None,
                         "growth_total_pct": None, "growth_nonbranded_pct": None})
            continue
        pts = _pick_trend_points(series)
        cur = pts["current"]
        ya = pts["year_ago"]

        def pct(now, then):
            if not then or then == 0:
                return None
            return round((now - then) / then * 100, 1)

        growth_total = pct(cur["total"], ya["total"]) if (cur and ya) else None
        growth_nb = pct(cur["non_branded"], ya["non_branded"]) if (cur and ya) else None

        print(f"[7]   {dom}: atual={cur['date']} total={cur['total']:,} "
              f"(brand={cur['branded']:,} / non-brand={cur['non_branded']:,})"
              + (f"  vs {ya['date']} total={ya['total']:,} (cresc {growth_total}%)" if ya else ""))

        if cur: cur_dates.add(cur["date"])
        if ya: ya_dates.add(ya["date"])
        rows.append({
            "domain": dom, "role": role,
            "perfil": profile_by_domain.get(dom, "—"),
            "current": cur, "year_ago": ya,
            "growth_total_pct": growth_total,
            "growth_nonbranded_pct": growth_nb,
        })

    return {
        "rows": rows,
        "current_date": ", ".join(sorted(cur_dates)) if cur_dates else "",
        "year_ago_date": ", ".join(sorted(ya_dates)) if ya_dates else "",
    }


def _score_competitor(domain: str, info: dict, briefing: dict) -> float:
    """Score de quao especializado o concorrente eh no nicho do cliente.
    Maior = melhor candidato para backup Semrush."""
    score = info["count"] * 10.0  # aparicoes na SERP pesam mais

    raiz = [t.lower() for t in (briefing.get("termos_raiz") or [])]
    nicho_principal = (briefing.get("nicho_principal") or "").lower()
    nicho_principal_norm = unicodedata.normalize("NFKD", nicho_principal).encode("ascii", "ignore").decode("ascii")
    nicho_words = [w for w in re.split(r"\W+", nicho_principal_norm) if len(w) >= 4]

    dom = domain.lower()
    is_subdomain = dom.count(".") >= 2
    if is_subdomain:
        score += 5  # subdominio dedicado (ex: saude.petlove.com.br) > root multi-produto
    for w in raiz + nicho_words:
        if w and w in dom:
            score += 4
            break

    # bonus se o nicho identificado pelo LLM bate com nicho do cliente
    for r in info["rows"]:
        n = (r.get("nicho") or "").lower()
        n_norm = unicodedata.normalize("NFKD", n).encode("ascii", "ignore").decode("ascii")
        for w in nicho_words:
            if w in n_norm:
                score += 2
                break

    # penalidade para dominios de marca muito conhecida (proxy: nome longo monolitico de conglomerado)
    brand = brand_token_from_domain(domain)
    if len(brand) >= 11 and brand not in dom.split(".")[0]:
        pass
    # heuristica anti-conglomerado: se o nome do dominio root tem > 10 chars sem separador,
    # provavel marca-mae (portoseguro, construtoraitajai etc)
    root = dom.split(".")[0]
    if not is_subdomain and len(root) >= 11:
        score -= 3

    score += (info["vol"] or 0) / 10000.0  # tiebreaker minusculo
    return score


def rank_high_similarity_competitors(similarity_rows: list, briefing: dict) -> list:
    """Retorna lista ordenada de candidatos Alto (melhor primeiro)."""
    alto_rows = [s for s in similarity_rows if s.get("similaridade") == "Alto" and s.get("domain")]
    if not alto_rows:
        return []
    by_domain = {}
    for s in alto_rows:
        d = s["domain"]
        by_domain.setdefault(d, {"count": 0, "vol": 0, "rows": []})
        by_domain[d]["count"] += 1
        by_domain[d]["vol"] += s.get("volume") or 0
        by_domain[d]["rows"].append(s)

    ranked = []
    for domain, info in by_domain.items():
        info["domain"] = domain
        info["score"] = _score_competitor(domain, info, briefing)
        ranked.append(info)
    ranked.sort(key=lambda x: x["score"], reverse=True)
    return ranked


def brand_token_from_domain(domain: str) -> str:
    """Extrai o token principal de marca do dominio (sem TLD/www)."""
    base = normalize_domain(domain).split(".")[0]
    base = re.sub(r"[^a-z0-9]+", "", base.lower())
    return base


def is_brand_keyword(phrase_norm: str, brand: str,
                      url_hints: list = None,
                      fuzzy_threshold: float = 0.82,
                      generic_terms: list = None) -> bool:
    """Considera marca se:
    1) brand aparece como substring na phrase (sem espacos), OU
    2) alguma palavra da phrase tem similaridade fuzzy >= threshold com a marca, OU
    3) o nome aparece em URL como /empreendimento/<algo>/ (nome proprio de produto da marca)

    `generic_terms`: termos genericos do nicho (ex.: termos_raiz). Se a palavra que
    casaria com a marca por fuzzy for, na verdade, um termo generico do nicho (ex.:
    marca 'leilo' x palavra 'leilao'), o match fuzzy e ignorado para evitar descartar
    keywords legitimas. O match exato por substring continua valendo.
    """
    if not brand:
        return False

    def _ascii(s: str) -> str:
        return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii").lower()

    brand_a = _ascii(brand)
    generic_norm = [_ascii(t) for t in (generic_terms or []) if t]

    def _is_generic(w: str) -> bool:
        """True se a palavra for um termo generico do nicho ou sua variacao
        (plural/flexao), ex.: 'leilao'/'leiloes' quando termo_raiz e 'leilao'."""
        for g in generic_norm:
            if not g:
                continue
            if SequenceMatcher(None, w, g).ratio() >= 0.75:
                return True
            common = 0
            for a, b in zip(w, g):
                if a == b:
                    common += 1
                else:
                    break
            if common >= 4 and abs(len(w) - len(g)) <= 3:
                return True
        return False

    words = [_ascii(w) for w in phrase_norm.split()]

    # 1) palavra identica a marca -> sempre marca (ex.: 'leilo', 'leilo app')
    if brand_a in words:
        return True

    # 2) substring/fuzzy por palavra, ignorando termos genericos do nicho
    #    (evita 'leilo' x 'leilao'/'leiloes' marcarem keywords de categoria)
    for w in words:
        if len(w) < 4:
            continue
        hit = (brand_a in w) or (SequenceMatcher(None, w, brand_a).ratio() >= fuzzy_threshold)
        if hit and not _is_generic(w):
            return True

    # 3) marca multi-palavra colada (ex.: 'chat guru' -> 'chatguru'),
    #    desde que nao seja explicada por uma unica palavra (ja tratada acima)
    compact = "".join(words)
    if len(brand_a) >= 5 and brand_a in compact and not any(brand_a in w for w in words):
        return True

    if url_hints:
        for u in url_hints:
            if not u:
                continue
            m = re.search(r"/empreendimento/([^/?#]+)", u.lower())
            if m:
                slug = re.sub(r"[^a-z0-9]+", " ", m.group(1)).strip()
                slug_compact = slug.replace(" ", "")
                if SequenceMatcher(None, compact, slug_compact).ratio() >= 0.7:
                    return True
    return False


def filter_backup_keywords(semrush_kws: list, existing_keywords: list,
                            competitor_domain: str) -> list:
    """Marca cada keyword com status NOVA / JA_NO_RELATORIO / MARCA."""
    existing = set()
    for k in existing_keywords:
        kw = k["keyword"]
        kw = unicodedata.normalize("NFKD", kw).encode("ascii", "ignore").decode("ascii").lower()
        kw = re.sub(r"\s+", " ", kw).strip()
        existing.add(kw)

    brand = brand_token_from_domain(competitor_domain)

    out = []
    for k in semrush_kws:
        phrase = (k.get("phrase") or "").lower()
        phrase_norm = unicodedata.normalize("NFKD", phrase).encode("ascii", "ignore").decode("ascii")
        phrase_norm = re.sub(r"\s+", " ", phrase_norm).strip()
        url = k.get("url") or ""
        if is_brand_keyword(phrase_norm, brand, url_hints=[url]):
            status = "MARCA"
        elif phrase_norm in existing:
            status = "JA NO RELATORIO"
        else:
            status = "NOVA"
        out.append({**k, "status": status, "phrase_norm": phrase_norm})
    return out


def run_top5_similarity(keywords: list, volumes: dict, serps: dict,
                        client_url: str, briefing: dict, top_kws: int = 5,
                        top_results: int = 3, workers: int = 4) -> list:
    candidates = sorted(
        [k for k in keywords if (volumes.get(k["keyword"], {}).get("volume_medio") or 0) > 0],
        key=lambda k: volumes[k["keyword"]]["volume_medio"] or 0,
        reverse=True,
    )

    if not candidates:
        print("[5] Sem keywords com volume para analise de similaridade.")
        return []

    sorted_kws, dropped = dedupe_keywords_by_intent(candidates, top_kws)
    if dropped:
        print(f"[5] {len(dropped)} variacao(oes) descartada(s) por intencao duplicada:")
        for d in dropped[:5]:
            print(f"      - '{d['keyword']}' (mesma intencao de '{d['duplicate_of']}')")
    print(f"[5] Top {len(sorted_kws)} keywords distintas por intencao: "
          + ", ".join([f"'{k['keyword']}'" for k in sorted_kws]))

    pairs = []
    for k in sorted_kws:
        items = serps.get(k["keyword"], [])[:top_results]
        for it in items:
            pairs.append({
                "keyword": k["keyword"],
                "volume": volumes[k["keyword"]]["volume_medio"],
                "rank": it["rank"],
                "domain": it["domain"],
                "url": it["url"],
                "title": it["title"],
            })

    unique_urls = list({p["url"] for p in pairs if p["url"]})
    print(f"[5] Analisando similaridade de {len(unique_urls)} URLs unicas ({len(pairs)} posicoes) ...")

    sim_cache = {}
    with ThreadPoolExecutor(max_workers=workers) as pool:
        future_to_url = {pool.submit(analyze_similarity, u, client_url, briefing): u for u in unique_urls}
        done = 0
        for fut in as_completed(future_to_url):
            url = future_to_url[fut]
            try:
                sim_cache[url] = fut.result()
            except Exception as e:
                sim_cache[url] = {"nicho": "—", "similaridade": "—", "perfil": "—", "motivo": f"Erro: {e}"}
            done += 1
            r = sim_cache[url]
            print(f"[5]   ({done}/{len(unique_urls)}) {r.get('similaridade','—')}/{r.get('perfil','—')}  {url[:60]}")

    results = []
    for p in pairs:
        s = sim_cache.get(p["url"], {})
        results.append({**p,
                        "nicho": s.get("nicho", ""),
                        "similaridade": s.get("similaridade", "—"),
                        "perfil": s.get("perfil", "—"),
                        "motivo": s.get("motivo", "")})
    return results


# ---------------------------------------------------------------------------
# SAIDA: XLSX
# ---------------------------------------------------------------------------

def write_xlsx(out_path: Path, url: str, briefing: dict, keywords: list,
               volumes: dict, positions: dict, top_concorrentes: list,
               similarity_rows: list, backup_info: dict,
               traffic_comparison: dict = None):
    print(f"[OUT] Escrevendo XLSX: {out_path}")
    wb = Workbook()

    # ---- Aba Briefing
    ws = wb.active
    ws.title = "Briefing"
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(["Campo", "Resposta"])
    for c in ws[1]:
        c.font = head_font
        c.fill = head_fill
        c.border = border
        c.alignment = Alignment(horizontal="center")

    ws.append(["URL", url])
    ws.append(["Data", datetime.now().strftime("%d/%m/%Y %H:%M")])
    for key, label in BRIEFING_FIELDS:
        ws.append([label, str(briefing.get(key, ""))])
    ws.append(["Cidades alvo", ", ".join(briefing.get("cidades") or [])])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 90

    # ---- Aba Personas & Canais
    personas = briefing.get("personas") or []
    canais = briefing.get("canais") or []
    if personas or canais:
        wsp = wb.create_sheet("Personas & Canais")
        if personas:
            wsp.append(["Personas"])
            wsp.cell(row=wsp.max_row, column=1).font = Font(bold=True, size=13, color="0F172A")
            wsp.append(["Persona", "Idade / Perfil", "Motivacoes", "Onde busca informacao"])
            for c in wsp[wsp.max_row]:
                c.font = head_font
                c.fill = head_fill
                c.border = border
                c.alignment = Alignment(horizontal="center")
            for p in personas:
                wsp.append([
                    p.get("nome", ""),
                    p.get("idade_perfil", ""),
                    p.get("descricao", ""),
                    p.get("onde_busca", ""),
                ])
            wsp.append([""])

        if canais:
            wsp.append(["Potencial de Canais de Marketing"])
            wsp.cell(row=wsp.max_row, column=1).font = Font(bold=True, size=13, color="0F172A")
            wsp.append(["Canal", "Conversao", "Velocidade", "Custo / Esforco", "Justificativa"])
            for c in wsp[wsp.max_row]:
                c.font = head_font
                c.fill = head_fill
                c.border = border
                c.alignment = Alignment(horizontal="center")

            fill_alta = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
            fill_media = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
            fill_baixa = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

            def _fill_for(v: str):
                s = (v or "").strip().lower()
                if s.startswith("alt"):
                    return fill_alta
                if s.startswith("med") or s.startswith("méd"):
                    return fill_media
                if s.startswith("bai"):
                    return fill_baixa
                return None

            for c in canais:
                wsp.append([
                    c.get("canal", ""),
                    c.get("conversao", ""),
                    c.get("velocidade", ""),
                    c.get("custo", ""),
                    c.get("justificativa", ""),
                ])
                row_idx = wsp.max_row
                for col_idx, val in zip((2, 3, 4),
                                         (c.get("conversao"), c.get("velocidade"), c.get("custo"))):
                    f = _fill_for(val)
                    if f is not None:
                        wsp.cell(row=row_idx, column=col_idx).fill = f
                    wsp.cell(row=row_idx, column=col_idx).alignment = Alignment(horizontal="center")

        for row in wsp.iter_rows(min_row=1, max_row=wsp.max_row):
            for cell in row:
                if cell.value not in (None, ""):
                    cell.alignment = Alignment(wrap_text=True, vertical="top",
                                                horizontal=cell.alignment.horizontal or "left")
        widths = [34, 38, 50, 18, 70]
        for i, w in enumerate(widths, 1):
            wsp.column_dimensions[chr(64 + i)].width = w

    # ---- Aba Palavras-chave
    ws2 = wb.create_sheet("Palavras-chave")
    headers = ["#", "Palavra-chave", "Funil", "Categoria",
               "Volume Medio", "Posicao Google", "Concorrencia"]
    ws2.append(headers)
    for c in ws2[1]:
        c.font = head_font
        c.fill = head_fill
        c.border = border
        c.alignment = Alignment(horizontal="center")

    sorted_kws = sorted(
        keywords,
        key=lambda k: (volumes.get(k["keyword"], {}).get("volume_medio") or 0),
        reverse=True,
    )

    for i, k in enumerate(sorted_kws, 1):
        v = volumes.get(k["keyword"], {})
        kw = k["keyword"]
        if v.get("volume_medio"):
            pos = positions.get(kw)
            pos_cell = pos if pos else "N/R"
        else:
            pos_cell = "—"
        ws2.append([
            i,
            kw,
            k.get("funil", ""),
            k.get("categoria", ""),
            v.get("volume_medio"),
            pos_cell,
            v.get("competition"),
        ])

    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row):
        for cell in row:
            cell.border = border

    widths = [5, 45, 10, 22, 14, 14, 16]
    for i, w in enumerate(widths, 1):
        ws2.column_dimensions[chr(64 + i)].width = w

    # ---- Aba Top 3 Concorrentes
    ws_top = wb.create_sheet("Top 3 Concorrentes")
    ws_top.append(["Palavra-chave", "Volume Medio", "Posicao", "Dominio", "URL", "Titulo"])
    for c in ws_top[1]:
        c.font = head_font
        c.fill = head_fill
        c.border = border
        c.alignment = Alignment(horizontal="center")
    for entry in top_concorrentes:
        kw = entry["keyword"]
        vol = entry["volume"]
        for r in entry["top3"]:
            ws_top.append([kw, vol, r["rank"], r["domain"], r["url"], r["title"]])
    for row in ws_top.iter_rows(min_row=2, max_row=ws_top.max_row):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate([35, 14, 10, 30, 60, 60], 1):
        ws_top.column_dimensions[chr(64 + i)].width = w

    # ---- Aba Similaridade Top 5
    ws_sim = wb.create_sheet("Similaridade Top 5")
    sim_headers = ["Keyword", "Volume", "Posicao", "Dominio", "URL", "Nicho",
                   "Similaridade", "Perfil", "Motivo"]
    ws_sim.append(sim_headers)
    for c in ws_sim[1]:
        c.font = head_font
        c.fill = head_fill
        c.border = border
        c.alignment = Alignment(horizontal="center")

    fills_sim = {
        "Alto":  PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid"),
        "Medio": PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
        "Baixo": PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"),
    }
    fills_perfil = {
        "Especialista": PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid"),
        "Generalista":  PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
    }

    for row in similarity_rows:
        ws_sim.append([
            row.get("keyword", ""),
            row.get("volume"),
            row.get("rank"),
            row.get("domain", ""),
            row.get("url", ""),
            row.get("nicho", ""),
            row.get("similaridade", ""),
            row.get("perfil", ""),
            row.get("motivo", ""),
        ])
        last = ws_sim.max_row
        fill = fills_sim.get(row.get("similaridade", ""))
        if fill:
            ws_sim.cell(row=last, column=7).fill = fill
        fill_p = fills_perfil.get(row.get("perfil", ""))
        if fill_p:
            ws_sim.cell(row=last, column=8).fill = fill_p
        for c in ws_sim[last]:
            c.border = border
            c.alignment = Alignment(wrap_text=True, vertical="top")

    for i, w in enumerate([28, 10, 10, 26, 55, 28, 14, 14, 70], 1):
        ws_sim.column_dimensions[chr(64 + i)].width = w

    # ---- Aba Backup Semrush (top 30 de 1 concorrente Alto)
    ws_bk = wb.create_sheet("Backup Semrush")
    if backup_info and backup_info.get("keywords"):
        comp = backup_info.get("competitor", "")
        ws_bk.append([f"Concorrente analisado: {comp}", "", "", "", "", "", "", ""])
        ws_bk["A1"].font = Font(bold=True, size=12)
        ws_bk.append(["#", "Keyword", "Status", "Posicao Concorrente",
                      "Volume", "KD%", "Trafego Estimado", "URL"])
        for c in ws_bk[2]:
            c.font = head_font
            c.fill = head_fill
            c.border = border
            c.alignment = Alignment(horizontal="center")

        fills_status = {
            "NOVA":             PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid"),
            "JA NO RELATORIO":  PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid"),
            "MARCA":            PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"),
        }
        for idx, k in enumerate(backup_info["keywords"], 1):
            ws_bk.append([
                idx, k.get("phrase", ""), k.get("status", ""),
                k.get("position"), k.get("volume"), k.get("kd"),
                k.get("traffic"), k.get("url", ""),
            ])
            last = ws_bk.max_row
            fill = fills_status.get(k.get("status", ""))
            if fill:
                ws_bk.cell(row=last, column=3).fill = fill
            for c in ws_bk[last]:
                c.border = border
                c.alignment = Alignment(vertical="top", wrap_text=True)
        for i, w in enumerate([5, 42, 18, 14, 12, 10, 16, 55], 1):
            ws_bk.column_dimensions[chr(64 + i)].width = w
    else:
        ws_bk.append(["Sem concorrente classificado como Alto na etapa de similaridade."])
        ws_bk["A1"].font = Font(italic=True, color="64748B")

    # ---- Aba Cliente vs Concorrentes (trafego SEO Semrush + perfil)
    if traffic_comparison and traffic_comparison.get("rows"):
        ws_tr = wb.create_sheet("Cliente vs Concorrentes")
        rows_tc = traffic_comparison["rows"]
        cur_label = traffic_comparison.get("current_date") or "Atual"
        ya_label = traffic_comparison.get("year_ago_date") or "12m atras"
        ws_tr.append(["Comparativo de trafego organico (Semrush) - cliente vs concorrentes Alto"])
        ws_tr.merge_cells(start_row=1, start_column=1, end_row=1, end_column=11)
        ws_tr["A1"].font = Font(bold=True, size=12, color="0F172A")
        ws_tr.append([])
        ws_tr.append([
            "Tipo", "Dominio", "Perfil",
            f"Atual ({cur_label}) - Non-Branded", f"Atual - Branded", f"Atual - Total",
            f"Ano antes ({ya_label}) - Non-Branded", "Ano antes - Branded", "Ano antes - Total",
            "Cresc. Non-Branded (%)", "Cresc. Total (%)",
        ])
        for c in ws_tr[3]:
            c.font = head_font
            c.fill = head_fill
            c.border = border
            c.alignment = Alignment(horizontal="center", wrap_text=True)

        client_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
        fills_perfil_tc = {
            "Especialista": PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"),
            "Generalista":  PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
        }
        for r in rows_tc:
            cur = r.get("current") or {}
            ya = r.get("year_ago") or {}
            ws_tr.append([
                r["role"].capitalize(),
                r["domain"],
                r.get("perfil", "—"),
                cur.get("non_branded"), cur.get("branded"), cur.get("total"),
                ya.get("non_branded"), ya.get("branded"), ya.get("total"),
                r.get("growth_nonbranded_pct"),
                r.get("growth_total_pct"),
            ])
            last = ws_tr.max_row
            for c in ws_tr[last]:
                c.border = border
                c.alignment = Alignment(vertical="center")
            if r["role"] == "cliente":
                for c in ws_tr[last]:
                    c.fill = client_fill
                    c.font = Font(bold=True)
            # destaque do perfil (so se o cliente NAO ja tem fill)
            fill_p = fills_perfil_tc.get(r.get("perfil", ""))
            if fill_p and r["role"] != "cliente":
                ws_tr.cell(row=last, column=3).fill = fill_p
            for col in (4, 5, 6, 7, 8, 9):
                ws_tr.cell(row=last, column=col).number_format = "#,##0"
            for col in (10, 11):
                v = ws_tr.cell(row=last, column=col).value
                ws_tr.cell(row=last, column=col).number_format = "+0.0\\%;-0.0\\%;0"
                if isinstance(v, (int, float)):
                    color = "16A34A" if v >= 0 else "DC2626"
                    ws_tr.cell(row=last, column=col).font = Font(
                        bold=(r["role"] == "cliente"), color=color)

        ws_tr.column_dimensions["A"].width = 14
        ws_tr.column_dimensions["B"].width = 32
        ws_tr.column_dimensions["C"].width = 14
        for col_letter in ["D", "E", "F", "G", "H", "I"]:
            ws_tr.column_dimensions[col_letter].width = 16
        ws_tr.column_dimensions["J"].width = 18
        ws_tr.column_dimensions["K"].width = 16

    # ---- Aba Resumo
    ws3 = wb.create_sheet("Resumo")
    total_vol = sum((volumes.get(k["keyword"], {}).get("volume_medio") or 0) for k in keywords)
    topo = sum(1 for k in keywords if k.get("funil") == "topo")
    fundo = sum(1 for k in keywords if k.get("funil") == "fundo")
    local = sum(1 for k in keywords if k.get("funil") == "local")
    kws_com_vol = [k for k in keywords if (volumes.get(k["keyword"], {}).get("volume_medio") or 0) > 0]
    ranqueadas = sum(1 for k in kws_com_vol if positions.get(k["keyword"]))
    top10 = sum(1 for k in kws_com_vol if (positions.get(k["keyword"]) or 999) <= 10)
    top3 = sum(1 for k in kws_com_vol if (positions.get(k["keyword"]) or 999) <= 3)
    ws3.append(["Metrica", "Valor"])
    ws3.append(["Total de keywords", len(keywords)])
    ws3.append(["Volume mensal total (soma)", total_vol])
    ws3.append(["Keywords com volume", len(kws_com_vol)])
    ws3.append(["Keywords ranqueando (top 20)", ranqueadas])
    ws3.append(["Keywords no Top 10", top10])
    ws3.append(["Keywords no Top 3", top3])
    ws3.append(["Keywords topo de funil", topo])
    ws3.append(["Keywords fundo de funil", fundo])
    ws3.append(["Keywords local", local])
    ws3.append(["Escopo", briefing.get("escopo", "")])
    if similarity_rows:
        sim_alto = sum(1 for s in similarity_rows if s.get("similaridade") == "Alto")
        sim_med = sum(1 for s in similarity_rows if s.get("similaridade") == "Medio")
        sim_bai = sum(1 for s in similarity_rows if s.get("similaridade") == "Baixo")
        ws3.append(["Similaridade Alto (top 5 x top 3)", sim_alto])
        ws3.append(["Similaridade Medio (top 5 x top 3)", sim_med])
        ws3.append(["Similaridade Baixo (top 5 x top 3)", sim_bai])
    if backup_info and backup_info.get("keywords"):
        novas = sum(1 for k in backup_info["keywords"] if k.get("status") == "NOVA")
        marcas = sum(1 for k in backup_info["keywords"] if k.get("status") == "MARCA")
        ja = sum(1 for k in backup_info["keywords"] if k.get("status") == "JA NO RELATORIO")
        ws3.append(["Backup Semrush - concorrente", backup_info.get("competitor", "")])
        ws3.append(["Backup Semrush - total", len(backup_info["keywords"])])
        ws3.append(["Backup Semrush - NOVAS", novas])
        ws3.append(["Backup Semrush - JA NO RELATORIO", ja])
        ws3.append(["Backup Semrush - MARCA (excluidas)", marcas])
    if traffic_comparison and traffic_comparison.get("rows"):
        cli = next((r for r in traffic_comparison["rows"] if r["role"] == "cliente"), None)
        if cli and cli.get("current"):
            ws3.append(["Trafego SEO - Total atual (cliente)", cli["current"]["total"]])
            ws3.append(["Trafego SEO - Non-Branded atual (cliente)", cli["current"]["non_branded"]])
            ws3.append(["Trafego SEO - Branded atual (cliente)", cli["current"]["branded"]])
            if cli.get("growth_total_pct") is not None:
                ws3.append(["Trafego SEO - Crescimento Total (%) 12m", cli["growth_total_pct"]])
            if cli.get("growth_nonbranded_pct") is not None:
                ws3.append(["Trafego SEO - Crescimento Non-Branded (%) 12m", cli["growth_nonbranded_pct"]])
    for c in ws3[1]:
        c.font = head_font
        c.fill = head_fill
    ws3.column_dimensions["A"].width = 38
    ws3.column_dimensions["B"].width = 24

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
    print(f"[OUT] XLSX salvo: {saved}")
    return saved


# ---------------------------------------------------------------------------
# SAIDA: HTML
# ---------------------------------------------------------------------------

HTML_TEMPLATE = """<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Relatorio SEO - {domain}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,sans-serif;background:#f8fafc;color:#1e293b;line-height:1.55}}
.wrap{{max-width:1100px;margin:0 auto;padding:40px 24px}}
.brand-bar{{height:5px;background:linear-gradient(90deg,#0f766e,#2563eb,#7c3aed);border-radius:4px;margin-bottom:32px}}
header{{margin-bottom:28px}}
header h1{{font-size:28px;font-weight:800;color:#0f172a;margin-bottom:6px}}
header .sub{{font-size:14px;color:#64748b;margin-bottom:14px}}
.info-row{{display:flex;flex-wrap:wrap;gap:8px 22px;font-size:13px;color:#475569}}
.info-row b{{color:#0f172a}}
.info-row a{{color:#2563eb;text-decoration:none}}
.niche-banner{{background:linear-gradient(135deg,#7c3aed,#2563eb);color:#fff;border-radius:12px;padding:24px 28px;margin-bottom:28px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:16px;box-shadow:0 4px 12px rgba(124,58,237,.18)}}
.niche-banner .niche-label{{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:1px;opacity:.85;margin-bottom:6px}}
.niche-banner .niche-val{{font-size:38px;font-weight:800;line-height:1}}
.niche-banner .niche-sub{{font-size:13px;opacity:.9;margin-top:4px}}
.niche-banner .niche-icon{{font-size:48px;opacity:.4;font-weight:800}}
.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:28px}}
.card{{background:#fff;border-radius:12px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.06);border-left:4px solid transparent}}
.card.c1{{border-left-color:#2563eb}}.card.c2{{border-left-color:#0f766e}}
.card.c3{{border-left-color:#f97316}}.card.c4{{border-left-color:#7c3aed}}
.card .label{{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:#64748b;margin-bottom:4px}}
.card .val{{font-size:28px;font-weight:800;color:#0f172a}}
.card .val.teal{{color:#0f766e}}.card .val.orange{{color:#f97316}}.card .val.purple{{color:#7c3aed}}
section{{margin-bottom:28px}}
section h2{{font-size:17px;font-weight:700;color:#0f172a;margin-bottom:12px;padding-bottom:8px;border-bottom:2px solid #e2e8f0}}
.brief-grid{{display:grid;grid-template-columns:1fr 1fr;gap:8px 24px;background:#fff;border-radius:10px;padding:18px;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.brief-grid .lab{{font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;font-weight:600;padding-top:8px}}
.brief-grid .val{{font-size:14px;color:#1e293b;padding-bottom:8px;border-bottom:1px solid #f1f5f9}}
.brief-grid .lab:last-of-type,.brief-grid .val:last-of-type{{border:none}}
.persona-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:14px}}
.persona-card{{background:#fff;border-radius:12px;padding:18px;box-shadow:0 1px 3px rgba(0,0,0,.06);border-top:4px solid #7c3aed}}
.persona-card h3{{font-size:15px;font-weight:800;color:#0f172a;margin-bottom:10px}}
.persona-card .pl{{font-size:10px;text-transform:uppercase;letter-spacing:.6px;color:#7c3aed;font-weight:700;margin-top:10px;margin-bottom:2px}}
.persona-card .pv{{font-size:13px;color:#334155;line-height:1.5}}
.canais-table{{width:100%;border-collapse:collapse;font-size:13px;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.04);margin-top:8px}}
.canais-table thead th{{background:#f1f5f9;color:#475569;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.5px;padding:12px 14px;text-align:left}}
.canais-table thead th.c{{text-align:center}}
.canais-table tbody td{{padding:10px 14px;border-top:1px solid #f1f5f9;color:#334155;vertical-align:middle}}
.canais-table tbody td.c{{text-align:center}}
.canais-table tbody tr:hover{{background:#f8fafc}}
.lv-Alta,.lv-Alto{{display:inline-block;background:rgba(16,185,129,.14);color:#047857;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:700;min-width:62px;text-align:center}}
.lv-Media,.lv-Medio{{display:inline-block;background:rgba(245,158,11,.14);color:#a16207;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:700;min-width:62px;text-align:center}}
.lv-Baixa,.lv-Baixo{{display:inline-block;background:rgba(220,38,38,.10);color:#b91c1c;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:700;min-width:62px;text-align:center}}
.canal-name{{font-weight:700;color:#0f172a}}
table{{width:100%;border-collapse:collapse;font-size:13px;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
thead th{{background:#f1f5f9;color:#475569;font-weight:700;font-size:11px;text-transform:uppercase;letter-spacing:.5px;padding:12px 14px;text-align:left}}
thead th.r{{text-align:right}}thead th.c{{text-align:center}}
tbody td{{padding:10px 14px;border-top:1px solid #f1f5f9;color:#334155}}
tbody td.r{{text-align:right;font-variant-numeric:tabular-nums}}tbody td.c{{text-align:center}}
tbody tr:hover{{background:#f8fafc}}
.vol-high{{color:#0f766e;font-weight:700}}
.tag{{display:inline-block;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:600;letter-spacing:.3px;text-transform:uppercase}}
.tag.topo{{background:rgba(2,132,199,.12);color:#0369a1}}
.tag.fundo{{background:rgba(220,38,38,.10);color:#b91c1c}}
.tag.local{{background:rgba(15,118,110,.12);color:#0f766e}}
.pos-top{{color:#0f766e;font-weight:700}}
.pos-mid{{color:#f97316;font-weight:700}}
.pos-low{{color:#dc2626;font-weight:700}}
.pos-na{{color:#94a3b8}}
.comp-section{{display:grid;grid-template-columns:1fr;gap:16px;margin-top:8px}}
.comp-card{{background:#fff;border-radius:10px;padding:18px;box-shadow:0 1px 3px rgba(0,0,0,.04);border-left:4px solid #7c3aed}}
.comp-card .ck{{font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:#7c3aed;font-weight:700;margin-bottom:2px}}
.comp-card .kw{{font-size:18px;font-weight:800;color:#0f172a;margin-bottom:2px}}
.comp-card .vol{{font-size:12px;color:#64748b;margin-bottom:12px}}
.comp-card ol{{padding-left:18px;margin:0}}
.comp-card li{{padding:8px 0;border-bottom:1px solid #f1f5f9}}
.comp-card li:last-child{{border-bottom:none}}
.comp-card li b{{color:#0f172a;font-size:14px}}
.comp-card li a{{color:#2563eb;text-decoration:none;font-size:12px;display:block;margin-top:2px;word-break:break-all}}
.comp-card li .title{{color:#475569;font-size:12px;margin-top:4px;line-height:1.4}}
.meta-tag{{display:inline-block;padding:3px 10px;background:#eef2ff;color:#3730a3;font-size:11px;font-weight:600;border-radius:6px;margin-left:8px}}
.tr-table{{margin-top:6px}}
.tr-table .role-cli{{background:#eff6ff;font-weight:700;color:#0f172a}}
.tr-table .num{{font-variant-numeric:tabular-nums;text-align:right}}
.tr-table .grow-up{{color:#16a34a;font-weight:700}}
.tr-table .grow-down{{color:#dc2626;font-weight:700}}
.tr-table .grow-na{{color:#94a3b8}}
.tr-table th.sub{{background:#e2e8f0;color:#334155;font-size:10px}}
.tr-legend{{font-size:12px;color:#64748b;margin-top:6px;margin-bottom:6px}}
.tr-legend b{{color:#0f172a}}
.perfil-tag{{display:inline-block;padding:3px 10px;border-radius:999px;font-size:11px;font-weight:600;letter-spacing:.3px}}
.perfil-tag.perf-esp{{background:rgba(34,197,94,.14);color:#15803d}}
.perfil-tag.perf-gen{{background:rgba(245,158,11,.14);color:#a16207}}
footer{{text-align:center;margin-top:32px;padding-top:20px;border-top:1px solid #e2e8f0;font-size:12px;color:#94a3b8}}
@media(max-width:760px){{.cards{{grid-template-columns:1fr 1fr}}.brief-grid{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<div class="wrap">
  <div class="brand-bar"></div>
  <header>
    <h1>Relatorio de Potencial SEO</h1>
    <div class="sub">{nicho} &middot; Escopo {escopo}{cidades_str}</div>
    <div class="info-row">
      <span>Cliente: <b><a href="{url}" target="_blank" rel="noopener">{domain}</a></b></span>
      <span>Base: <b>Brasil (br)</b></span>
      <span>Termos: <b>{n_kws}</b></span>
      <span>Volume total: <b>{vol_total}</b> /mes</span>
      <span>Data: <b>{date_str}</b></span>
    </div>
  </header>

  <div class="niche-banner">
    <div>
      <div class="niche-label">Volume Mensal Total do Nicho</div>
      <div class="niche-val">{vol_total}</div>
      <div class="niche-sub">somando todas as {n_kws} palavras analisadas</div>
    </div>
    <div class="niche-icon">{n_kws}</div>
  </div>

  <div class="cards">
    <div class="card c1"><div class="label">Palavras analisadas</div><div class="val">{n_kws}</div></div>
    <div class="card c2"><div class="label">Top 3 keywords</div><div class="val teal">{top3_vol}</div></div>
    <div class="card c3"><div class="label">Maior volume</div><div class="val orange">{max_vol}</div></div>
    <div class="card c4"><div class="label">{label_split}</div><div class="val purple">{split_text}</div></div>
  </div>

  <section>
    <h2>Briefing do Negocio</h2>
    <div class="brief-grid">
{brief_html}
    </div>
  </section>

{personas_section}

{canais_section}

  <section>
    <h2>Palavras-Chave &mdash; {n_kws} termos (ordenado por volume) <span class="meta-tag">DataForSEO &middot; Google Ads</span></h2>
    <table>
      <thead>
        <tr>
          <th>#</th><th>Palavra-chave</th><th class="c">Funil</th>
          <th class="r">Volume Medio /mes</th><th class="c">Posicao Google</th>
        </tr>
      </thead>
      <tbody>
{rows_html}
      </tbody>
    </table>
  </section>

  <section>
    <h2>Empresas no Top 3 &mdash; Top 2 Palavras de Maior Volume <span class="meta-tag">SERP DataForSEO</span></h2>
    <div class="comp-section">
{concorrentes_html}
    </div>
  </section>

{traffic_section}

  <footer>Gerado automaticamente &middot; {date_str} &middot; {domain}</footer>
</div>
</body>
</html>"""


def position_html(pos, has_volume: bool):
    if not has_volume:
        return '<span class="pos-na">—</span>'
    if pos is None:
        return '<span class="pos-low">N/R</span>'
    if pos <= 3:
        cls = "pos-top"
    elif pos <= 10:
        cls = "pos-mid"
    else:
        cls = "pos-low"
    return f'<span class="{cls}">#{pos}</span>'


def _traffic_growth_cell(v):
    if v is None:
        return '<td class="num grow-na">—</td>'
    cls = "grow-up" if v >= 0 else "grow-down"
    sign = "+" if v >= 0 else ""
    return f'<td class="num {cls}">{sign}{v:.1f}%</td>'


def _year_label(date_str: str) -> str:
    """'2026-04' -> '2026'. Multiplas datas viram a primeira."""
    if not date_str:
        return ""
    first = date_str.split(",")[0].strip()
    return first.split("-")[0] if "-" in first else first


def build_personas_section_html(personas: list) -> str:
    if not personas:
        return ""
    cards = []
    for p in personas:
        nome = (p.get("nome") or "Persona").strip()
        idade = (p.get("idade_perfil") or "").strip()
        desc = (p.get("descricao") or "").strip()
        onde = (p.get("onde_busca") or "").strip()
        cards.append(
            f'<div class="persona-card">'
            f'<h3>{nome}</h3>'
            + (f'<div class="pl">Idade / Perfil</div><div class="pv">{idade}</div>' if idade else "")
            + (f'<div class="pl">Motivacoes</div><div class="pv">{desc}</div>' if desc else "")
            + (f'<div class="pl">Onde busca informacao</div><div class="pv">{onde}</div>' if onde else "")
            + '</div>'
        )
    return (
        '  <section>\n'
        '    <h2>Publico-alvo &mdash; Personas <span class="meta-tag">Analise estrategica</span></h2>\n'
        '    <div class="persona-grid">\n'
        + "\n".join("      " + c for c in cards)
        + '\n    </div>\n'
        '  </section>'
    )


def _lv_cls(v: str) -> str:
    s = (v or "").strip()
    norm = s.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    return f"lv-{norm.capitalize()}" if norm else ""


def build_canais_section_html(canais: list) -> str:
    if not canais:
        return ""
    body = []
    for c in canais:
        canal = (c.get("canal") or "").strip()
        conv = (c.get("conversao") or "—").strip()
        vel = (c.get("velocidade") or "—").strip()
        custo = (c.get("custo") or "—").strip()
        just = (c.get("justificativa") or "").strip()
        body.append(
            "<tr>"
            f'<td><span class="canal-name">{canal}</span></td>'
            f'<td class="c"><span class="{_lv_cls(conv)}">{conv}</span></td>'
            f'<td class="c"><span class="{_lv_cls(vel)}">{vel}</span></td>'
            f'<td class="c"><span class="{_lv_cls(custo)}">{custo}</span></td>'
            f'<td>{just}</td>'
            "</tr>"
        )
    return (
        '  <section>\n'
        '    <h2>Projecao de Potencial por Canal de Marketing <span class="meta-tag">Analise estrategica</span></h2>\n'
        '    <table class="canais-table">\n'
        '      <thead><tr>'
        '<th>Canal de Marketing</th>'
        '<th class="c">Conversao</th>'
        '<th class="c">Velocidade de Retorno</th>'
        '<th class="c">Custo / Esforco</th>'
        '<th>Justificativa Estrategica</th>'
        '</tr></thead>\n'
        '      <tbody>\n        '
        + "\n        ".join(body)
        + '\n      </tbody>\n'
        '    </table>\n'
        '  </section>'
    )


def build_traffic_section_html(tc: dict) -> str:
    if not tc or not tc.get("rows"):
        return ""
    rows_tc = tc["rows"]
    cur_year = _year_label(tc.get("current_date", "")) or "Atual"
    ya_year = _year_label(tc.get("year_ago_date", "")) or "Ano antes"

    body = []
    for r in rows_tc:
        cur = r.get("current") or {}
        ya = r.get("year_ago") or {}
        cls = "role-cli" if r["role"] == "cliente" else ""
        role_label = "Cliente" if r["role"] == "cliente" else "Concorrente"
        perfil = r.get("perfil", "—")
        perfil_cls = ""
        if perfil == "Especialista":
            perfil_cls = "perf-esp"
        elif perfil == "Generalista":
            perfil_cls = "perf-gen"
        body.append(
            f'<tr class="{cls}">'
            f'<td>{role_label}</td>'
            f'<td><b>{r["domain"]}</b></td>'
            f'<td class="c"><span class="perfil-tag {perfil_cls}">{perfil}</span></td>'
            f'<td class="num"><b>{fmt_int(cur.get("non_branded"))}</b></td>'
            f'<td class="num" style="color:#64748b">{fmt_int(ya.get("non_branded"))}</td>'
            f'<td class="num">{fmt_int(cur.get("branded"))}</td>'
            f'<td class="num" style="color:#64748b">{fmt_int(ya.get("branded"))}</td>'
            f'{_traffic_growth_cell(r.get("growth_nonbranded_pct"))}'
            f'</tr>'
        )

    return f"""  <section>
    <h2>Cliente vs Concorrentes &mdash; Trafego SEO <span class="meta-tag">Semrush Organic Overview</span></h2>
    <div class="tr-legend"><b>SEO</b> = trafego de termos genericos (non-branded, o que importa para crescer). <b>Marca</b> = pesquisas pelo nome da empresa (branded). <b>Perfil</b>: Especialista (foco unico no nicho do cliente) ou Generalista (varios produtos). Comparativo {cur_year} vs {ya_year}.</div>
    <table class="tr-table">
      <thead>
        <tr>
          <th>Tipo</th>
          <th>Dominio</th>
          <th class="c">Perfil</th>
          <th class="c">SEO {cur_year}</th>
          <th class="c">SEO {ya_year}</th>
          <th class="c">Marca {cur_year}</th>
          <th class="c">Marca {ya_year}</th>
          <th class="c">Cresc. SEO</th>
        </tr>
      </thead>
      <tbody>
{"".join(body)}
      </tbody>
    </table>
  </section>"""


def write_html(out_path: Path, url: str, briefing: dict, keywords: list,
               volumes: dict, positions: dict, top_concorrentes: list,
               traffic_comparison: dict = None):
    print(f"[OUT] Escrevendo HTML: {out_path}")
    domain = urlparse(url).netloc or url
    sorted_kws = sorted(
        keywords,
        key=lambda k: (volumes.get(k["keyword"], {}).get("volume_medio") or 0),
        reverse=True,
    )

    vol_total = sum((volumes.get(k["keyword"], {}).get("volume_medio") or 0) for k in keywords)
    vols_sorted = [(volumes.get(k["keyword"], {}).get("volume_medio") or 0) for k in sorted_kws]
    top3_vol = sum(vols_sorted[:3])
    max_vol = vols_sorted[0] if vols_sorted else 0

    rows = []
    for i, k in enumerate(sorted_kws, 1):
        kw = k["keyword"]
        v = volumes.get(kw, {})
        vm = v.get("volume_medio")
        funil = k.get("funil", "")
        rows.append(
            f'<tr><td>{i}</td><td>{kw}</td>'
            f'<td class="c"><span class="tag {funil}">{funil}</span></td>'
            f'<td class="r vol-high">{fmt_int(vm)}</td>'
            f'<td class="c">{position_html(positions.get(kw), bool(vm))}</td></tr>'
        )

    conc_cards = []
    for entry in top_concorrentes:
        items_html = []
        for r in entry["top3"]:
            items_html.append(
                f'<li><b>#{r["rank"]} &middot; {r["domain"]}</b>'
                f'<a href="{r["url"]}" target="_blank" rel="noopener">{r["url"]}</a>'
                f'<div class="title">{r["title"]}</div></li>'
            )
        conc_cards.append(
            '<div class="comp-card">'
            f'<div class="ck">Palavra-chave</div>'
            f'<div class="kw">{entry["keyword"]}</div>'
            f'<div class="vol">{fmt_int(entry["volume"])} pesquisas/mes</div>'
            f'<ol>{"".join(items_html)}</ol>'
            '</div>'
        )
    concorrentes_html = "\n".join(conc_cards) if conc_cards else '<p style="color:#94a3b8">Sem dados de SERP disponiveis.</p>'

    escopo = briefing.get("escopo", "")
    cidades = briefing.get("cidades") or []
    cidades_str = f" &middot; {', '.join(cidades)}" if cidades else ""

    if escopo == "Local":
        label_split = "Cidades alvo"
        split_text = str(len(cidades))
    else:
        topo = sum(1 for k in keywords if k.get("funil") == "topo")
        fundo = sum(1 for k in keywords if k.get("funil") == "fundo")
        label_split = "Topo / Fundo"
        split_text = f"{topo}/{fundo}"

    brief_items = []
    for key, label in BRIEFING_FIELDS:
        val = str(briefing.get(key, "") or "—")
        brief_items.append(f'      <div class="lab">{label}</div><div class="val">{val}</div>')
    if cidades:
        brief_items.append(f'      <div class="lab">Cidades</div><div class="val">{", ".join(cidades)}</div>')

    personas_section = build_personas_section_html(briefing.get("personas") or [])
    canais_section = build_canais_section_html(briefing.get("canais") or [])

    traffic_section = build_traffic_section_html(traffic_comparison)

    html = HTML_TEMPLATE.format(
        domain=domain,
        url=url,
        nicho=briefing.get("nicho", "") or "—",
        escopo=escopo,
        cidades_str=cidades_str,
        n_kws=len(keywords),
        vol_total=fmt_int(vol_total),
        top3_vol=fmt_int(top3_vol),
        max_vol=fmt_int(max_vol),
        label_split=label_split,
        split_text=split_text,
        date_str=datetime.now().strftime("%d/%m/%Y"),
        brief_html="\n".join(brief_items),
        personas_section=personas_section,
        canais_section=canais_section,
        rows_html="\n".join(rows),
        concorrentes_html=concorrentes_html,
        traffic_section=traffic_section,
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html, encoding="utf-8")
    print(f"[OUT] HTML salvo: {out_path}")
    return out_path


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def run(url: str):
    t0 = time.time()
    slug = domain_slug(url)
    out_dir = OUTPUT_DIR / slug
    print(f"\n=== Pipeline SEO: {url} ===")
    print(f"Output dir: {out_dir}\n")

    site = scrape_site(url)
    if not site.get("body"):
        print("[!] Site vazio. Abortando.")
        return

    briefing = llm_briefing(url, site)
    print(f"[1B] Briefing: escopo={briefing['escopo']}, nicho='{briefing.get('nicho_principal')}', "
          f"cidades={briefing.get('cidades')}, termos_raiz={briefing.get('termos_raiz')}")

    persona_canais = llm_personas_canais(url, site, briefing)
    briefing["personas"] = persona_canais.get("personas") or []
    briefing["canais"] = persona_canais.get("canais") or []

    candidates = llm_generate_keywords(briefing, site)
    if not candidates:
        print("[!] Nenhuma keyword gerada. Abortando.")
        return

    volumes = fetch_search_volume([k["keyword"] for k in candidates])

    client_brand = brand_token_from_domain(urlparse(url).netloc)
    keywords = filter_keywords_by_volume(candidates, volumes, briefing["escopo"],
                                          brand=client_brand,
                                          generic_terms=briefing.get("termos_raiz"))

    kws_with_volume = [k for k in keywords
                       if (volumes.get(k["keyword"], {}).get("volume_medio") or 0) > 0]
    print(f"[3] {len(kws_with_volume)}/{len(keywords)} keywords no relatorio final com volume > 0")

    client_domain = urlparse(url).netloc
    serps = fetch_serps([k["keyword"] for k in kws_with_volume])

    positions = {}
    for k in kws_with_volume:
        positions[k["keyword"]] = find_position(serps.get(k["keyword"], []), client_domain)

    sorted_by_vol = sorted(
        kws_with_volume,
        key=lambda k: volumes[k["keyword"]]["volume_medio"] or 0,
        reverse=True,
    )
    top2_distintas, dropped_top2 = dedupe_keywords_by_intent(sorted_by_vol, 2)
    if dropped_top2:
        print(f"[4] {len(dropped_top2)} variacao(oes) descartada(s) na escolha do top 2:")
        for d in dropped_top2[:3]:
            print(f"      - '{d['keyword']}' (mesma intencao de '{d['duplicate_of']}')")

    top_concorrentes = []
    for k in top2_distintas:
        kw = k["keyword"]
        items = serps.get(kw, [])[:3]
        top_concorrentes.append({
            "keyword": kw,
            "volume": volumes[kw]["volume_medio"],
            "top3": items,
        })
    print(f"[4] Top concorrentes capturados para {len(top_concorrentes)} keywords distintas: "
          + ", ".join([f"'{e['keyword']}'" for e in top_concorrentes]))

    similarity_rows = run_top5_similarity(
        keywords, volumes, serps, client_url=url, briefing=briefing,
        top_kws=5, top_results=3, workers=4,
    )

    backup_info = {}
    candidates = rank_high_similarity_competitors(similarity_rows, briefing)
    if not candidates:
        print("[6] Nenhum concorrente Alto. Pulando backup Semrush.")
    else:
        print("[6] Candidatos Alto rankeados:")
        for c in candidates[:5]:
            print(f"      score={c['score']:.1f}  count={c['count']}  vol={c['vol']}  {c['domain']}")

        attempts = []
        for cand in candidates[:3]:
            comp_domain = cand["domain"]
            print(f"[6] Tentando concorrente: {comp_domain} (score {cand['score']:.1f}) ...")
            sem_kws = fetch_semrush_top_keywords(comp_domain, limit=SEMRUSH_TOP_N, db=SEMRUSH_DB)
            if not sem_kws:
                print(f"[6]   sem dados Semrush, proximo.")
                continue
            backup_kws = filter_backup_keywords(sem_kws, keywords, comp_domain)
            n_total = len(backup_kws)
            n_novas = sum(1 for k in backup_kws if k["status"] == "NOVA")
            n_marca = sum(1 for k in backup_kws if k["status"] == "MARCA")
            pct_marca = (n_marca / n_total) if n_total else 1
            print(f"[6]   {n_novas} novas / "
                  f"{sum(1 for k in backup_kws if k['status']=='JA NO RELATORIO')} ja / "
                  f"{n_marca} marca ({pct_marca*100:.0f}%)")
            attempts.append({
                "competitor": comp_domain, "keywords": backup_kws,
                "n_novas": n_novas, "pct_marca": pct_marca,
            })
            # criterio de aceitacao: pelo menos 50% nao-marca E >=10 novas
            if pct_marca <= 0.5 and n_novas >= 10:
                backup_info = {"competitor": comp_domain, "keywords": backup_kws}
                print(f"[6]   ACEITO: {comp_domain}")
                break
            print(f"[6]   muito marca, tentando proximo...")

        if not backup_info and attempts:
            # fallback: escolhe o melhor entre os tentados (maior n_novas)
            best = max(attempts, key=lambda a: a["n_novas"])
            backup_info = {"competitor": best["competitor"], "keywords": best["keywords"]}
            print(f"[6] Nenhum candidato passou o filtro - usando o com mais novas: {best['competitor']}")

    # ETAPA 7: comparativo de trafego SEO (cliente vs ate 3 concorrentes Alto)
    print("[7] Classificando perfil do cliente (Especialista/Generalista) ...")
    client_profile = classify_client_profile(url, site, briefing)
    print(f"[7] Cliente classificado como: {client_profile}")
    traffic_comparison = build_traffic_comparison(
        url, similarity_rows, max_competitors=3, client_profile=client_profile,
    )

    xlsx_path = out_dir / f"relatorio-seo-{slug}.xlsx"
    html_path = out_dir / f"relatorio-seo-{slug}.html"
    write_xlsx(xlsx_path, url, briefing, keywords, volumes, positions, top_concorrentes,
               similarity_rows, backup_info, traffic_comparison)
    write_html(html_path, url, briefing, keywords, volumes, positions, top_concorrentes,
               traffic_comparison)

    print(f"\n=== Concluido em {time.time()-t0:.1f}s ===")
    print(f"XLSX: {xlsx_path}")
    print(f"HTML: {html_path}")


def main():
    if len(sys.argv) < 2:
        print("Uso: python seo_pipeline.py <url>")
        print("Ex:  python seo_pipeline.py https://cnempreendimentos.net/")
        sys.exit(1)
    url = sys.argv[1].strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    run(url)


if __name__ == "__main__":
    main()
