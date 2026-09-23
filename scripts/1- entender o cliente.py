# -*- coding: utf-8 -*-
"""
Pipeline de coleta de dados do cliente + briefing + personas/canais.

Extrai do seo_pipeline.py apenas as etapas:
  [1A] Scraping HTML
  [1B] LLM Briefing (14 campos)
  [1C] LLM Personas & Canais

Saidas:
  - XLSX: abas Briefing + Personas & Canais
  - HTML: relatorio visual (briefing + personas + canais)

Uso:
    python cliente_briefing_pipeline.py
    python cliente_briefing_pipeline.py input.txt
    python cliente_briefing_pipeline.py https://chatguru.com.br/
"""

import os
import re
import sys
import json
import time
import importlib.util
import unicodedata
from pathlib import Path
from urllib.parse import urlparse
from datetime import datetime

import requests
import openpyxl
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "").strip()
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
OPENROUTER_MODEL = "google/gemini-2.5-flash"

BASE_DIR = Path(__file__).resolve().parent
# Outputs centralizados no Radar 09 2026
try:
    from workspace_paths import PIPELINE_OUTPUT_DIR as OUTPUT_DIR, FORMULA_DIR  # noqa: E402
except Exception:
    OUTPUT_DIR = BASE_DIR.parent / "outputs" / "entender"
    FORMULA_DIR = Path(r"C:\Users\Usuario\Desktop\Radar Concorrencia\Formula de potencial de canais")
DEFAULT_INPUT_FILE = BASE_DIR / "input.txt"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


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


def normalize_label(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "").encode("ascii", "ignore").decode("ascii")
    value = value.lower().replace("?", "")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


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

CLASSIFICAR2_FIELDS = [
    ("nicho_amplo", "Nicho Amplo"),
    ("nicho_medio", "Nicho Medio"),
    ("nicho_especifico", "Nicho Especifico"),
    ("produto_principal", "Produto principal"),
    ("produtos", "Produtos (homepage)"),
    ("perfil", "Perfil"),
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


def llm_classificar2(url: str, site: dict, briefing: dict) -> dict:
    print("[1B2] Gerando Classificar 2 (homepage + perfil) ...")
    prompt = f"""Leia SOMENTE a homepage abaixo e classifique o negocio em niveis de nicho, produto principal, produtos visiveis e perfil.

URL: {url}

CONTEUDO DA HOMEPAGE:
{site_signals_text(site)[:5000]}

CONTEXTO DO BRIEFING ATUAL:
- Nicho: {briefing.get('nicho', '')}
- Nicho principal: {briefing.get('nicho_principal', '')}
- Nicho secundario: {briefing.get('nicho_secundario', '')}
- Tipo de negocio: {briefing.get('tipo_negocio', '')}
- Modelo de produto: {briefing.get('modelo_produto', '')}

RESPONDA APENAS JSON VALIDO com estas chaves:
- "nicho_amplo": setor/macro, 1-2 palavras. Ex: Saude, Tecnologia, Educacao, Imoveis
- "nicho_medio": categoria, 1-3 palavras. Ex: Medicina, Servicos de TI, Odontologia
- "nicho_especifico": especialidade visivel na homepage, 1-4 palavras
- "produto_principal": oferta de maior destaque na homepage (hero, H1 ou primeiro bloco comercial)
- "produtos": string com 4 a 10 ofertas/servicos visiveis na homepage, separadas por virgula
- "perfil": "Especialista" ou "Generalista"

REGRAS:
- Nao use nome da marca nos campos de nicho.
- Nicho Amplo > Nicho Medio > Nicho Especifico devem formar um funil coerente.
- Extraia PRODUTOS e PRODUTO PRINCIPAL SOMENTE do que aparece na homepage (title, menu, H1, H2, cards, blocos comerciais).
- "perfil" = Especialista se o foco unico ou principal da homepage estiver no mesmo nicho principal do briefing.
- "perfil" = Generalista se a homepage oferecer multiplos produtos/servicos e o nicho principal do briefing for apenas um entre varios.

Responda somente o JSON, sem texto antes ou depois."""
    raw = call_openrouter(prompt, max_tokens=700, temperature=0.1)
    data = extract_json(raw)
    for key, _ in CLASSIFICAR2_FIELDS:
        data.setdefault(key, "")
    return data


def as_list(value) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    if isinstance(value, str):
        parts = re.split(r"\n|;|,", value)
        return [p.strip(" -•\t") for p in parts if p.strip(" -•\t")]
    return [str(value)]


def empty_traducao() -> dict:
    return {
        "nicho": {"amplo": "—", "medio": "—", "especifico": "—"},
        "produto": {
            "amplo": "—", "medio": "—", "especifico": "—",
            "exemplo_dor": "—", "antes": "—", "depois": "—", "beneficios": "—",
        },
        "produtos": [],
        "dores_exemplos": [],
        "publico_alvo": "—",
        "audiencia_empresas": [],
        "audiencia_cargos": [],
        "dores_que_resolve": [],
        "motivacoes": [],
        "onde_busca_informacao": [],
    }


def load_formula_engine():
    engine_path = FORMULA_DIR / "engine.py"
    if not engine_path.exists():
        raise FileNotFoundError(f"engine.py nao encontrado em {engine_path}")
    spec = importlib.util.spec_from_file_location("formula_engine", engine_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def llm_formula_canais(url: str, briefing: dict) -> dict:
    print("[1E] Gerando Formula de Potencial de Canais ...")
    engine = load_formula_engine()
    motores_validos = sorted(engine.MOTORES.keys())
    contexto_opcoes = {
        "idade": ["jovem", "adulto", "maduro", "senior"],
        "modelo": ["B2B", "B2C"],
        "alcance": ["local", "regional", "nacional"],
        "ticket": ["baixo", "medio", "alto"],
        "visualidade": ["baixa", "media", "alta"],
        "confianca_risco": ["baixo", "medio", "alto", "muito_alto"],
        "surpresa": ["baixa", "media", "alta"],
        "dor_desejo": ["dor", "misto", "desejo"],
        "potencial_organico": ["baixo", "medio", "alto"],
    }
    traducao = briefing.get("traducao") or {}
    prompt = f"""Voce e um estrategista de canais. A partir do briefing e da traducao da homepage, preencha os insumos da Formula de Potencial de Canais.

URL: {url}

BRIEFING:
- Nicho: {briefing.get('nicho', '')}
- Nicho principal: {briefing.get('nicho_principal', '')}
- Nicho secundario: {briefing.get('nicho_secundario', '')}
- Publico alvo: {briefing.get('publico_alvo', '')}
- B2B/B2C: {briefing.get('b2b_b2c', '')}
- Escopo: {briefing.get('escopo', '')}
- Tipo de negocio: {briefing.get('tipo_negocio', '')}
- Modelo de produto: {briefing.get('modelo_produto', '')}
- Dores: {briefing.get('dores', '')}

TRADUCAO:
- Nicho amplo: {(traducao.get('nicho') or {}).get('amplo', '')}
- Nicho medio: {(traducao.get('nicho') or {}).get('medio', '')}
- Nicho especifico: {(traducao.get('nicho') or {}).get('especifico', '')}
- Produto principal: {(traducao.get('produto') or {}).get('especifico', '')}
- Publico alvo traduzido: {traducao.get('publico_alvo', '')}
- Dores que resolve: {', '.join(traducao.get('dores_que_resolve') or [])}
- Motivacoes: {', '.join(traducao.get('motivacoes') or [])}

RETORNE SOMENTE JSON VALIDO neste formato:
{{
  "produto": {{
    "nicho": "",
    "empresa": "",
    "produto": "",
    "compra_real": "",
    "emocao_antes": "",
    "emocao_depois": "",
    "observacao": ""
  }},
  "motores": {{
    "motor": 0
  }},
  "contexto": {{
    "idade": "",
    "modelo": "",
    "alcance": "",
    "ticket": "",
    "visualidade": "",
    "confianca_risco": "",
    "surpresa": "",
    "dor_desejo": "",
    "potencial_organico": ""
  }}
}}

REGRAS:
- Em "motores", retorne de 5 a 8 motores com peso de 0 a 10.
- Use somente estes motores validos: {", ".join(motores_validos)}.
- "modelo" deve ser exatamente uma das opcoes: B2B, B2C.
- "alcance" deve ser exatamente uma das opcoes: local, regional, nacional.
- Para cada campo de contexto, use somente uma destas opcoes:
  idade={contexto_opcoes["idade"]}
  ticket={contexto_opcoes["ticket"]}
  visualidade={contexto_opcoes["visualidade"]}
  confianca_risco={contexto_opcoes["confianca_risco"]}
  surpresa={contexto_opcoes["surpresa"]}
  dor_desejo={contexto_opcoes["dor_desejo"]}
  potencial_organico={contexto_opcoes["potencial_organico"]}
- "produto.nicho" deve ser um nicho curto e generico.
- "produto.produto" deve ser a oferta principal mais clara.
- "compra_real" responde o que o cliente realmente compra, nao a feature.
- "emocao_antes" e "emocao_depois" devem ser curtas.
- "observacao" resume o tipo de demanda e o racional de canais.
"""
    raw = call_openrouter(prompt, max_tokens=1400, temperature=0.2)
    data = extract_json(raw)

    produto = data.get("produto") or {}
    motores = {}
    for motor, peso in (data.get("motores") or {}).items():
        if motor in engine.MOTORES:
            try:
                peso_n = int(round(float(peso)))
            except Exception:
                continue
            motores[motor] = max(0, min(10, peso_n))

    contexto = {}
    for chave, opcoes in contexto_opcoes.items():
        val = str((data.get("contexto") or {}).get(chave) or "").strip()
        if val in opcoes:
            contexto[chave] = val

    # fallbacks para contexto essencial
    contexto.setdefault("modelo", "B2B" if str(briefing.get("b2b_b2c", "")).upper().startswith("B2B") else "B2C")
    contexto.setdefault("alcance", "nacional" if str(briefing.get("escopo", "")).lower().startswith("nac") else "local")
    contexto.setdefault("idade", "adulto")
    contexto.setdefault("ticket", "medio")
    contexto.setdefault("visualidade", "media")
    contexto.setdefault("confianca_risco", "alto")
    contexto.setdefault("surpresa", "media")
    contexto.setdefault("dor_desejo", "dor")
    contexto.setdefault("potencial_organico", "medio")

    if not motores:
        motores = {"confianca": 9, "conveniencia": 8, "escala": 7, "performance": 7, "economia": 6}

    produto_obj = engine.Produto(
        nicho=str(produto.get("nicho") or briefing.get("nicho_principal") or briefing.get("nicho") or "—"),
        empresa=urlparse(url).netloc or url,
        produto=str(produto.get("produto") or briefing.get("produto_principal") or "Oferta principal"),
        compra_real=str(produto.get("compra_real") or "Eficiência, escala e organização da operação"),
        emocao_antes=str(produto.get("emocao_antes") or "Desorganização"),
        emocao_depois=str(produto.get("emocao_depois") or "Controle e confiança"),
        motores=motores,
        contexto=contexto,
        observacao=str(produto.get("observacao") or ""),
    )
    res = engine.analisar(produto_obj)
    return {
        "produto": {
            "nicho": produto_obj.nicho,
            "empresa": produto_obj.empresa,
            "produto": produto_obj.produto,
            "compra_real": produto_obj.compra_real,
            "emocao_antes": produto_obj.emocao_antes,
            "emocao_depois": produto_obj.emocao_depois,
            "observacao": produto_obj.observacao,
        },
        "motores": motores,
        "contexto": contexto,
        "top_motores": res.get("top_motores") or [],
        "ranking_potencial": res.get("ranking_potencial") or [],
        "ranking_mercado": res.get("ranking_mercado") or [],
        "ranking_oportunidade": res.get("ranking_oportunidade") or [],
        "estrategia": res.get("estrategia") or {},
    }


def llm_traducao_homepage(url: str, site: dict, briefing: dict) -> dict:
    print("[1D] Gerando Traducao da homepage ...")
    prompt = f"""Voce e um estrategista de posicionamento. Sua tarefa e TRADUZIR a homepage:
decifrar o jargao corporativo e reescrever em linguagem comercial clara, util para filtros e copy.

URL: {url}

HOMEPAGE:
{site_signals_text(site)[:5000]}

BRIEFING JA EXISTENTE (use como contexto, mas priorize a homepage):
- Nicho: {briefing.get('nicho', '')}
- Nicho principal: {briefing.get('nicho_principal', '')}
- Nicho secundario: {briefing.get('nicho_secundario', '')}
- Publico alvo: {briefing.get('publico_alvo', '')}
- Modelo de produto: {briefing.get('modelo_produto', '')}

Retorne SOMENTE JSON valido neste formato:
{{
  "nicho": {{
    "amplo": "setor/macro, 1-2 palavras",
    "medio": "categoria, 1-3 palavras",
    "especifico": "especialidade visivel na homepage, 1-4 palavras"
  }},
  "produto": {{
    "amplo": "familia do que vendem",
    "medio": "tipo de oferta",
    "especifico": "oferta de maior destaque no hero/H1/menu",
    "exemplo_dor": "1 cenario concreto de dor do comprador",
    "antes": "como a vida/operacao do cliente e antes de contratar",
    "depois": "como fica depois de contratar",
    "beneficios": ["1 a 3 tags. Use SOMENTE: Dinheiro, Tempo, Qualidade, Risco, Conformidade"]
  }},
  "produtos": [
    {{
      "nome": "nome da oferta na homepage",
      "amplo": "",
      "medio": "",
      "especifico": "",
      "exemplo_dor": "",
      "antes": "",
      "depois": "",
      "beneficios": ["Dinheiro"]
    }}
  ],
  "dores_exemplos": ["4 a 8 exemplos concretos de dor"],
  "publico_alvo": "1-2 frases descrevendo quem compra",
  "audiencia_empresas": ["6 a 10 tipos de empresa"],
  "audiencia_cargos": ["6 a 10 cargos tomadores de decisao"],
  "dores_que_resolve": ["4 a 8 dores que o negocio realmente resolve"],
  "motivacoes": ["4 a 6 motivacoes de compra"],
  "onde_busca_informacao": ["4 a 8 canais/fontes onde esse publico busca informacao"]
}}

REGRAS:
- Use SOMENTE o que da para inferir da PAGINA INICIAL (title, menu, H1, H2, cards). Nao invente paginas internas.
- Nicho e o MERCADO. Produto e o que a empresa VENDE. Nao misture os dois.
- Funil obrigatorio: especifico cabe no medio, medio cabe no amplo.
- Nao use nome da marca nos 3 niveis de nicho.
- "produto" descreve a oferta PRINCIPAL da homepage.
- Em "produtos", retorne no maximo 2 itens: o produto principal e mais 1 produto complementar visivel, se existir.
- Antes/depois em linguagem do cliente, nao slogan.
- "beneficios": use somente Dinheiro, Tempo, Qualidade, Risco, Conformidade.
- Listas curtas, sem repeticao, portugues do Brasil.
"""
    raw = call_openrouter(prompt, max_tokens=2600, temperature=0.2)
    data = extract_json(raw)
    out = empty_traducao()

    nicho = data.get("nicho") or {}
    out["nicho"] = {
        "amplo": str(nicho.get("amplo") or "—"),
        "medio": str(nicho.get("medio") or "—"),
        "especifico": str(nicho.get("especifico") or "—"),
    }

    prod = data.get("produto") or {}
    out["produto"] = {
        "amplo": str(prod.get("amplo") or "—"),
        "medio": str(prod.get("medio") or "—"),
        "especifico": str(prod.get("especifico") or "—"),
        "exemplo_dor": str(prod.get("exemplo_dor") or "—"),
        "antes": str(prod.get("antes") or "—"),
        "depois": str(prod.get("depois") or "—"),
        "beneficios": ", ".join(as_list(prod.get("beneficios"))) or "—",
    }

    produtos = []
    for item in data.get("produtos") or []:
        if not isinstance(item, dict):
            continue
        produtos.append({
            "nome": str(item.get("nome") or "—"),
            "amplo": str(item.get("amplo") or "—"),
            "medio": str(item.get("medio") or "—"),
            "especifico": str(item.get("especifico") or "—"),
            "exemplo_dor": str(item.get("exemplo_dor") or "—"),
            "antes": str(item.get("antes") or "—"),
            "depois": str(item.get("depois") or "—"),
            "beneficios": ", ".join(as_list(item.get("beneficios"))) or "—",
        })
        if len(produtos) >= 2:
            break
    if not produtos and out["produto"].get("especifico", "—") != "—":
        produtos.append({
            "nome": out["produto"]["especifico"],
            "amplo": out["produto"]["amplo"],
            "medio": out["produto"]["medio"],
            "especifico": out["produto"]["especifico"],
            "exemplo_dor": out["produto"]["exemplo_dor"],
            "antes": out["produto"]["antes"],
            "depois": out["produto"]["depois"],
            "beneficios": out["produto"]["beneficios"],
        })
    out["produtos"] = produtos
    out["dores_exemplos"] = as_list(data.get("dores_exemplos"))
    out["publico_alvo"] = str(data.get("publico_alvo") or "—")
    out["audiencia_empresas"] = as_list(data.get("audiencia_empresas"))
    out["audiencia_cargos"] = as_list(data.get("audiencia_cargos"))
    out["dores_que_resolve"] = as_list(data.get("dores_que_resolve"))
    out["motivacoes"] = as_list(data.get("motivacoes"))
    out["onde_busca_informacao"] = as_list(data.get("onde_busca_informacao"))
    return out


# ---------------------------------------------------------------------------
# SAIDA: XLSX
# ---------------------------------------------------------------------------

def write_xlsx(out_path: Path, url: str, briefing: dict):
    print(f"[OUT] Escrevendo XLSX: {out_path}")
    wb = Workbook()

    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---- Aba Briefing
    ws = wb.active
    ws.title = "Briefing"
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
    for key, label in CLASSIFICAR2_FIELDS:
        ws.append([label, str(briefing.get(key, ""))])
    ws.append(["Cidades alvo", ", ".join(briefing.get("cidades") or [])])
    ws.append(["Termos raiz", ", ".join(briefing.get("termos_raiz") or [])])

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

    traducao = briefing.get("traducao") or {}
    if traducao:
        nicho = traducao.get("nicho") or {}
        prod = traducao.get("produto") or {}
        wst = wb.create_sheet("Traducao")
        wst.append(["Campo", "Resposta"])
        for c in wst[1]:
            c.font = head_font
            c.fill = head_fill
            c.border = border
            c.alignment = Alignment(horizontal="center")
        rows = [
            ("Nicho Amplo", nicho.get("amplo", "")),
            ("Nicho Medio", nicho.get("medio", "")),
            ("Nicho Especifico", nicho.get("especifico", "")),
            ("Produto Amplo", prod.get("amplo", "")),
            ("Produto Medio", prod.get("medio", "")),
            ("Produto Especifico", prod.get("especifico", "")),
            ("Produto exemplo dor", prod.get("exemplo_dor", "")),
            ("Produto antes", prod.get("antes", "")),
            ("Produto depois", prod.get("depois", "")),
            ("Produto beneficios", prod.get("beneficios", "")),
            ("Publico alvo", traducao.get("publico_alvo", "")),
            ("Audiencia Empresas", ", ".join(traducao.get("audiencia_empresas") or [])),
            ("Audiencia Cargos", ", ".join(traducao.get("audiencia_cargos") or [])),
            ("Dores exemplos", " | ".join(traducao.get("dores_exemplos") or [])),
            ("Dores que resolve", " | ".join(traducao.get("dores_que_resolve") or [])),
            ("Motivacoes", " | ".join(traducao.get("motivacoes") or [])),
            ("Onde busca informacao", " | ".join(traducao.get("onde_busca_informacao") or [])),
        ]
        for item in rows:
            wst.append(list(item))
        for row in wst.iter_rows(min_row=2, max_row=wst.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        wst.column_dimensions["A"].width = 30
        wst.column_dimensions["B"].width = 95

        wsp2 = wb.create_sheet("Traducao Produtos")
        wsp2.append(["Nome", "Amplo", "Medio", "Especifico", "Exemplo dor", "Antes", "Depois", "Beneficios"])
        for c in wsp2[1]:
            c.font = head_font
            c.fill = head_fill
            c.border = border
            c.alignment = Alignment(horizontal="center")
        for p in traducao.get("produtos") or []:
            wsp2.append([
                p.get("nome", ""), p.get("amplo", ""), p.get("medio", ""), p.get("especifico", ""),
                p.get("exemplo_dor", ""), p.get("antes", ""), p.get("depois", ""), p.get("beneficios", ""),
            ])
        for row in wsp2.iter_rows(min_row=2, max_row=wsp2.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        widths = [26, 22, 24, 28, 38, 42, 42, 24]
        for i, w in enumerate(widths, 1):
            wsp2.column_dimensions[chr(64 + i)].width = w

    formula = briefing.get("formula_canais") or {}
    if formula:
        wsf = wb.create_sheet("Formula Canais")
        wsf.append(["Campo", "Resposta"])
        for c in wsf[1]:
            c.font = head_font
            c.fill = head_fill
            c.border = border
            c.alignment = Alignment(horizontal="center")
        prod = formula.get("produto") or {}
        rows = [
            ("Produto / nicho", f"{prod.get('produto', '')} | {prod.get('nicho', '')}"),
            ("Compra real", prod.get("compra_real", "")),
            ("Emocao antes", prod.get("emocao_antes", "")),
            ("Emocao depois", prod.get("emocao_depois", "")),
            ("Observacao", prod.get("observacao", "")),
            ("Motores", ", ".join(f"{k}:{v}" for k, v in (formula.get("top_motores") or []))),
            ("Contexto", ", ".join(f"{k}={v}" for k, v in (formula.get("contexto") or {}).items())),
            ("Top Potencial", " | ".join(f"{c['canal']} ({c['potencial']})" for c in (formula.get("ranking_potencial") or [])[:5])),
            ("Top Oportunidade", " | ".join(f"{c['canal']} ({c['oportunidade']:+g})" for c in (formula.get("ranking_oportunidade") or [])[:5])),
            ("Curto prazo", ", ".join(c["canal"] for c in (formula.get("estrategia") or {}).get("curto", []))),
            ("Medio prazo", ", ".join(c["canal"] for c in (formula.get("estrategia") or {}).get("medio", []))),
            ("Longo prazo", ", ".join(c["canal"] for c in (formula.get("estrategia") or {}).get("longo", []))),
        ]
        for item in rows:
            wsf.append(list(item))
        for row in wsf.iter_rows(min_row=2, max_row=wsf.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        wsf.column_dimensions["A"].width = 26
        wsf.column_dimensions["B"].width = 100

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
<title>Briefing do Cliente - {domain}</title>
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
.niche-banner .niche-val{{font-size:32px;font-weight:800;line-height:1}}
.niche-banner .niche-sub{{font-size:13px;opacity:.9;margin-top:4px}}
.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:28px}}
.card{{background:#fff;border-radius:12px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,.06);border-left:4px solid transparent}}
.card.c1{{border-left-color:#2563eb}}.card.c2{{border-left-color:#0f766e}}
.card.c3{{border-left-color:#f97316}}.card.c4{{border-left-color:#7c3aed}}
.card .label{{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:#64748b;margin-bottom:4px}}
.card .val{{font-size:22px;font-weight:800;color:#0f172a}}
.card .val.teal{{color:#0f766e}}.card .val.orange{{color:#f97316}}.card .val.purple{{color:#7c3aed}}
.tabs{{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 22px;padding-bottom:10px;border-bottom:2px solid #e2e8f0}}
.tab-btn{{border:0;background:#e2e8f0;color:#334155;padding:9px 14px;border-radius:999px;font-size:12px;font-weight:700;cursor:pointer}}
.tab-btn.active{{background:#2563eb;color:#fff}}
.tab-panel{{display:none}}
.tab-panel.active{{display:block}}
section{{margin-bottom:28px}}
section h2{{font-size:17px;font-weight:700;color:#0f172a;margin-bottom:12px;padding-bottom:8px;border-bottom:2px solid #e2e8f0}}
.section-stack{{display:grid;gap:20px}}
.brief-block{{background:#fff;border-radius:12px;padding:18px 20px;box-shadow:0 1px 3px rgba(0,0,0,.04)}}
.brief-block h3{{font-size:14px;font-weight:800;color:#0f766e;margin-bottom:14px}}
.brief-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px 24px}}
.brief-item{{padding-bottom:12px;border-bottom:1px solid #f1f5f9}}
.brief-item.wide{{grid-column:1 / -1}}
.brief-grid .lab{{font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#64748b;font-weight:700;margin-bottom:5px}}
.brief-grid .val{{font-size:14px;color:#1e293b}}
.translate-levels{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px}}
.translate-level{{background:#fff;border-radius:12px;padding:18px;box-shadow:0 1px 3px rgba(0,0,0,.06);border-top:4px solid #7c3aed}}
.translate-level.l2{{border-top-color:#2563eb}}.translate-level.l3{{border-top-color:#0f766e}}
.translate-level .lab{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:#64748b;margin-bottom:6px}}
.translate-level .val{{font-size:20px;font-weight:800;color:#0f172a}}
.translate-story{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}}
.translate-card{{background:#fff;border-radius:12px;padding:18px;box-shadow:0 1px 3px rgba(0,0,0,.06)}}
.translate-card.dor{{border-left:4px solid #dc2626}}
.translate-card.antes{{border-left:4px solid #d97706}}
.translate-card.depois{{border-left:4px solid #059669}}
.translate-card.beneficio{{border-left:4px solid #2563eb}}
.translate-card .lab{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:#64748b;margin-bottom:8px}}
.translate-card .val{{font-size:14px;color:#334155}}
.chips{{display:flex;flex-wrap:wrap;gap:8px}}
.chip{{display:inline-block;background:#eef2ff;color:#3730a3;padding:6px 12px;border-radius:999px;font-size:13px;font-weight:600}}
.chip.red{{background:#fef2f2;color:#b91c1c}}
.chip.teal{{background:#ecfdf5;color:#047857}}
.chip.amber{{background:#fffbeb;color:#b45309}}
.block{{background:#fff;border-radius:12px;padding:18px;box-shadow:0 1px 3px rgba(0,0,0,.04);margin-bottom:12px}}
.block .lab{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#64748b;margin-bottom:8px}}
.block .val{{font-size:15px;color:#1e293b}}
.grid2{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
.empty,.empty-inline{{font-size:13px;color:#64748b}}
.empty{{background:#fff;padding:14px 16px;border-radius:10px}}
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
.meta-tag{{display:inline-block;padding:3px 10px;background:#eef2ff;color:#3730a3;font-size:11px;font-weight:600;border-radius:6px;margin-left:8px}}
footer{{text-align:center;margin-top:32px;padding-top:20px;border-top:1px solid #e2e8f0;font-size:12px;color:#94a3b8}}
@media(max-width:760px){{.cards{{grid-template-columns:1fr 1fr}}.brief-grid{{grid-template-columns:1fr}}.brief-item.wide{{grid-column:auto}}.translate-levels,.translate-story,.grid2{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<div class="wrap">
  <div class="brand-bar"></div>
  <header>
    <h1>Briefing do Cliente</h1>
    <div class="sub">{nicho} &middot; Escopo {escopo}{cidades_str}</div>
    <div class="info-row">
      <span>Cliente: <b><a href="{url}" target="_blank" rel="noopener">{domain}</a></b></span>
      <span>Modelo: <b>{modelo_produto}</b></span>
      <span>Data: <b>{date_str}</b></span>
    </div>
  </header>

  <div class="niche-banner">
    <div>
      <div class="niche-label">Nicho Principal</div>
      <div class="niche-val">{nicho_principal}</div>
      <div class="niche-sub">{b2b_b2c} &middot; {tipo_negocio}</div>
    </div>
  </div>

  <div class="cards">
    <div class="card c1"><div class="label">Escopo</div><div class="val">{escopo}</div></div>
    <div class="card c2"><div class="label">Personas</div><div class="val teal">{n_personas}</div></div>
    <div class="card c3"><div class="label">Canais analisados</div><div class="val orange">{n_canais}</div></div>
    <div class="card c4"><div class="label">B2B / B2C</div><div class="val purple">{b2b_b2c}</div></div>
  </div>

  <nav class="tabs">
    <button type="button" class="tab-btn active" data-tab="principal">Entender o Cliente</button>
    <button type="button" class="tab-btn" data-tab="extra">Aba Extra</button>
    <button type="button" class="tab-btn" data-tab="traducao">Traducao</button>
    <button type="button" class="tab-btn" data-tab="formula">Formula</button>
  </nav>

  <div class="tab-panel active" id="panel-principal">
{principal_section}
  </div>

  <div class="tab-panel" id="panel-extra">
{extra_section}
  </div>

  <div class="tab-panel" id="panel-traducao">
{traducao_section}
  </div>

  <div class="tab-panel" id="panel-formula">
{formula_section}
  </div>

  <footer>Gerado automaticamente &middot; {date_str} &middot; {domain}</footer>
</div>
<script>
document.querySelectorAll('.tab-btn').forEach(function(btn) {{
  btn.addEventListener('click', function() {{
    document.querySelectorAll('.tab-btn').forEach(function(el) {{ el.classList.remove('active'); }});
    document.querySelectorAll('.tab-panel').forEach(function(el) {{ el.classList.remove('active'); }});
    btn.classList.add('active');
    document.getElementById('panel-' + btn.dataset.tab).classList.add('active');
  }});
}});
</script>
</body>
</html>"""


def _lv_cls(v: str) -> str:
    s = (v or "").strip()
    norm = s.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    return f"lv-{norm.capitalize()}" if norm else ""


def _brief_item_html(label: str, value: str, wide: bool = False) -> str:
    cls = "brief-item wide" if wide else "brief-item"
    return f'        <div class="{cls}"><div class="lab">{label}</div><div class="val">{value}</div></div>'


def _chips_html(items: list[str], cls: str = "chip") -> str:
    if not items:
        return '<span class="empty-inline">—</span>'
    return "".join(f'<span class="{cls}">{item}</span>' for item in items)


def build_principal_section_html(briefing: dict) -> str:
    bloco_negocio = [
        _brief_item_html("Nicho", briefing.get("nicho", "") or "—"),
        _brief_item_html("Publico Alvo", briefing.get("publico_alvo", "") or "—"),
        _brief_item_html("Exemplo audiencia", briefing.get("audiencia_alvos", "") or "—", wide=True),
        _brief_item_html("B2B ou B2C", briefing.get("b2b_b2c", "") or "—"),
        _brief_item_html("Local ou Nacional", briefing.get("escopo", "") or "—"),
        _brief_item_html("Servico, Digital, Ecommerce", briefing.get("tipo_negocio", "") or "—"),
        _brief_item_html("Modelo de produto", briefing.get("modelo_produto", "") or "—"),
        _brief_item_html("Dores que resolve", briefing.get("dores", "") or "—", wide=True),
    ]
    bloco_classificar = [
        _brief_item_html("Nicho Principal", briefing.get("nicho_principal", "") or "—"),
        _brief_item_html("Nicho Secundario", briefing.get("nicho_secundario", "") or "—"),
        _brief_item_html("Palavras para SEO", briefing.get("seo_keywords", "") or "—"),
        _brief_item_html("Termos Raiz", ", ".join(briefing.get("termos_raiz") or []) or "—", wide=True),
    ]
    bloco_classificar2 = [
        _brief_item_html("Nicho Principal", briefing.get("nicho_principal", "") or "—"),
        _brief_item_html("Nicho Amplo", briefing.get("nicho_amplo", "") or "—"),
        _brief_item_html("Nicho Medio", briefing.get("nicho_medio", "") or "—"),
        _brief_item_html("Nicho Especifico", briefing.get("nicho_especifico", "") or "—"),
        _brief_item_html("Produto principal", briefing.get("produto_principal", "") or "—"),
        _brief_item_html("Produtos (homepage)", briefing.get("produtos", "") or "—", wide=True),
        _brief_item_html("Perfil", briefing.get("perfil", "") or "—"),
    ]
    return (
        '  <section>\n'
        '    <h2>Entender o Cliente</h2>\n'
        '    <div class="section-stack">\n'
        '      <div class="brief-block">\n'
        '        <h3>Briefing do Negocio</h3>\n'
        '        <div class="brief-grid">\n'
        + "\n".join(bloco_negocio)
        + '\n        </div>\n'
        '      </div>\n'
        '      <div class="brief-block">\n'
        '        <h3>Classificar</h3>\n'
        '        <div class="brief-grid">\n'
        + "\n".join(bloco_classificar)
        + '\n        </div>\n'
        '      </div>\n'
        '      <div class="brief-block">\n'
        '        <h3>Classificar 2</h3>\n'
        '        <div class="brief-grid">\n'
        + "\n".join(bloco_classificar2)
        + '\n        </div>\n'
        '      </div>\n'
        '    </div>\n'
        '  </section>'
    )


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


def build_extra_section_html(briefing: dict) -> str:
    extras = [
        _brief_item_html("Ferramenta/SaaS", briefing.get("ferramenta_saas", "") or "—"),
        _brief_item_html("Curso", briefing.get("curso", "") or "—"),
        _brief_item_html("Dores que Resolve", briefing.get("dores", "") or "—", wide=True),
        _brief_item_html("AI", briefing.get("ai", "") or "—"),
    ]
    return (
        '  <section>\n'
        '    <h2>Aba Extra</h2>\n'
        '    <div class="section-stack">\n'
        '      <div class="brief-block">\n'
        '        <h3>Complementos do Negocio</h3>\n'
        '        <div class="brief-grid">\n'
        + "\n".join(extras)
        + '\n        </div>\n'
        '      </div>\n'
        '    </div>\n'
        '  </section>\n'
        + build_personas_section_html(briefing.get("personas") or [])
        + '\n\n'
        + build_canais_section_html(briefing.get("canais") or [])
    )


def build_traducao_section_html(briefing: dict) -> str:
    traducao = briefing.get("traducao") or {}
    nicho = traducao.get("nicho") or {}
    prod = traducao.get("produto") or {}
    produtos = traducao.get("produtos") or []

    prod_rows = []
    for p in produtos:
        prod_rows.append(
            "<tr>"
            f"<td><b>{p.get('nome', '—')}</b></td>"
            f"<td>{p.get('amplo', '—')}</td>"
            f"<td>{p.get('medio', '—')}</td>"
            f"<td>{p.get('especifico', '—')}</td>"
            f"<td>{p.get('exemplo_dor', '—')}</td>"
            f"<td>{p.get('antes', '—')}</td>"
            f"<td>{p.get('depois', '—')}</td>"
            f"<td>{_chips_html(as_list(p.get('beneficios')), 'chip teal')}</td>"
            "</tr>"
        )
    prod_table = (
        '<p class="empty">Nenhum produto extraido da homepage.</p>'
        if not prod_rows else
        '<table class="canais-table"><thead><tr>'
        '<th>Produto</th><th>Amplo</th><th>Medio</th><th>Especifico</th>'
        '<th>Exemplo dor</th><th>Antes</th><th>Depois</th><th>Beneficios</th>'
        '</tr></thead><tbody>'
        + "".join(prod_rows)
        + "</tbody></table>"
    )

    return (
        '  <section>\n'
        '    <h2>Traducao</h2>\n'
        '    <div class="brief-block">\n'
        '      <h3>Nicho</h3>\n'
        '      <div class="translate-levels">\n'
        f'        <div class="translate-level"><div class="lab">Amplo</div><div class="val">{nicho.get("amplo", "—")}</div></div>\n'
        f'        <div class="translate-level l2"><div class="lab">Medio</div><div class="val">{nicho.get("medio", "—")}</div></div>\n'
        f'        <div class="translate-level l3"><div class="lab">Especifico</div><div class="val">{nicho.get("especifico", "—")}</div></div>\n'
        '      </div>\n'
        '    </div>\n'
        '  </section>\n'
        '  <section>\n'
        '    <h2>Produto principal <span class="meta-tag">Amplo / Medio / Especifico + dor</span></h2>\n'
        '    <div class="translate-levels">\n'
        f'      <div class="translate-level"><div class="lab">Amplo</div><div class="val">{prod.get("amplo", "—")}</div></div>\n'
        f'      <div class="translate-level l2"><div class="lab">Medio</div><div class="val">{prod.get("medio", "—")}</div></div>\n'
        f'      <div class="translate-level l3"><div class="lab">Especifico</div><div class="val">{prod.get("especifico", "—")}</div></div>\n'
        '    </div>\n'
        '    <div class="translate-story">\n'
        f'      <div class="translate-card dor"><div class="lab">Exemplo de dor</div><div class="val">{prod.get("exemplo_dor", "—")}</div></div>\n'
        f'      <div class="translate-card antes"><div class="lab">Antes</div><div class="val">{prod.get("antes", "—")}</div></div>\n'
        f'      <div class="translate-card depois"><div class="lab">Depois</div><div class="val">{prod.get("depois", "—")}</div></div>\n'
        f'      <div class="translate-card beneficio"><div class="lab">Beneficios</div><div class="val">{_chips_html(as_list(prod.get("beneficios")), "chip teal")}</div></div>\n'
        '    </div>\n'
        '  </section>\n'
        '  <section>\n'
        '    <h2>Produtos da homepage</h2>\n'
        f'    {prod_table}\n'
        '  </section>\n'
        '  <section>\n'
        '    <h2>Dor — exemplos</h2>\n'
        f'    <div class="chips">{_chips_html(traducao.get("dores_exemplos") or [], "chip red")}</div>\n'
        '  </section>\n'
        '  <section>\n'
        '    <h2>Publico</h2>\n'
        f'    <div class="block"><div class="lab">Publico alvo</div><div class="val">{traducao.get("publico_alvo", "—")}</div></div>\n'
        '    <div class="grid2">\n'
        f'      <div class="block"><div class="lab">Audiencia Empresas</div><div class="chips">{_chips_html(traducao.get("audiencia_empresas") or [], "chip teal")}</div></div>\n'
        f'      <div class="block"><div class="lab">Audiencia Cargos</div><div class="chips">{_chips_html(traducao.get("audiencia_cargos") or [], "chip")}</div></div>\n'
        '    </div>\n'
        '  </section>\n'
        '  <section>\n'
        '    <h2>Dores que resolve</h2>\n'
        f'    <div class="chips">{_chips_html(traducao.get("dores_que_resolve") or [], "chip red")}</div>\n'
        '  </section>\n'
        '  <section>\n'
        '    <h2>Motivacoes</h2>\n'
        f'    <div class="chips">{_chips_html(traducao.get("motivacoes") or [], "chip amber")}</div>\n'
        '  </section>\n'
        '  <section>\n'
        '    <h2>Onde busca informacao</h2>\n'
        f'    <div class="chips">{_chips_html(traducao.get("onde_busca_informacao") or [], "chip")}</div>\n'
        '  </section>'
    )


def build_formula_section_html(briefing: dict) -> str:
    formula = briefing.get("formula_canais") or {}
    if not formula:
        return '  <section>\n    <h2>Formula de Potencial de Canais</h2>\n    <p class="empty">Sem dados da formula.</p>\n  </section>'

    prod = formula.get("produto") or {}
    top_motores = formula.get("top_motores") or []
    contexto = formula.get("contexto") or {}
    rank_pot = formula.get("ranking_potencial") or []
    rank_op = formula.get("ranking_oportunidade") or []
    estrategia = formula.get("estrategia") or {}

    motores_html = "\n".join(
        _brief_item_html(str(motor).replace("_", " ").title(), str(peso))
        for motor, peso in top_motores[:8]
    ) or _brief_item_html("Motores", "—")

    def _ranking_rows(items: list, score_key: str) -> str:
        rows = []
        for item in items[:6]:
            rows.append(
                "<tr>"
                f"<td><span class=\"canal-name\">{item.get('canal', '')}</span></td>"
                f"<td>{item.get('classe', '')}</td>"
                f"<td class=\"c\">{item.get(score_key, '')}</td>"
                "</tr>"
            )
        return "".join(rows) or "<tr><td colspan='3'>—</td></tr>"

    return (
        '  <section>\n'
        '    <h2>Formula de Potencial de Canais</h2>\n'
        '    <div class="section-stack">\n'
        '      <div class="brief-block">\n'
        '        <h3>Produto analisado</h3>\n'
        '        <div class="brief-grid">\n'
        f'{_brief_item_html("Nicho", prod.get("nicho", "—"))}\n'
        f'{_brief_item_html("Produto", prod.get("produto", "—"))}\n'
        f'{_brief_item_html("Compra real", prod.get("compra_real", "—"), wide=True)}\n'
        f'{_brief_item_html("Emocao antes", prod.get("emocao_antes", "—"))}\n'
        f'{_brief_item_html("Emocao depois", prod.get("emocao_depois", "—"))}\n'
        f'{_brief_item_html("Observacao", prod.get("observacao", "—"), wide=True)}\n'
        '        </div>\n'
        '      </div>\n'
        '      <div class="brief-block">\n'
        '        <h3>Motores principais</h3>\n'
        '        <div class="brief-grid">\n'
        f'{motores_html}\n'
        '        </div>\n'
        '      </div>\n'
        '      <div class="brief-block">\n'
        '        <h3>Contexto aplicado</h3>\n'
        f'        <div class="chips">{_chips_html([f"{k}: {v}" for k, v in contexto.items()], "chip")}</div>\n'
        '      </div>\n'
        '    </div>\n'
        '  </section>\n'
        '  <section>\n'
        '    <h2>Ranking de Potencial</h2>\n'
        '    <table class="canais-table">\n'
        '      <thead><tr><th>Canal</th><th>Classe</th><th class="c">Potencial</th></tr></thead>\n'
        f'      <tbody>{_ranking_rows(rank_pot, "potencial")}</tbody>\n'
        '    </table>\n'
        '  </section>\n'
        '  <section>\n'
        '    <h2>Ranking de Oportunidade</h2>\n'
        '    <table class="canais-table">\n'
        '      <thead><tr><th>Canal</th><th>Classe</th><th class="c">Oportunidade</th></tr></thead>\n'
        f'      <tbody>{_ranking_rows(rank_op, "oportunidade")}</tbody>\n'
        '    </table>\n'
        '  </section>\n'
        '  <section>\n'
        '    <h2>Estrategia por Horizonte</h2>\n'
        '    <div class="grid2">\n'
        f'      <div class="block"><div class="lab">Curto prazo</div><div class="val">{_chips_html([c.get("canal", "") for c in estrategia.get("curto", [])], "chip teal")}</div></div>\n'
        f'      <div class="block"><div class="lab">Medio prazo</div><div class="val">{_chips_html([c.get("canal", "") for c in estrategia.get("medio", [])], "chip teal")}</div></div>\n'
        f'      <div class="block"><div class="lab">Longo prazo</div><div class="val">{_chips_html([c.get("canal", "") for c in estrategia.get("longo", [])], "chip teal")}</div></div>\n'
        '    </div>\n'
        '  </section>'
    )


def read_briefing_bundle_from_xlsx(xlsx_path: Path) -> tuple[str, dict]:
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    fields = {}
    if "Briefing" in wb.sheetnames:
        ws = wb["Briefing"]
        for row in ws.iter_rows(values_only=True):
            if not row or row[0] is None:
                continue
            key = str(row[0]).strip()
            val = "" if len(row) < 2 or row[1] is None else str(row[1]).strip()
            fields[key] = val

    norm_fields = {normalize_label(k): v for k, v in fields.items()}

    def get_field(*labels: str, default: str = "") -> str:
        for label in labels:
            val = norm_fields.get(normalize_label(label))
            if val:
                return val
        return default

    def get_by_prefix(*prefixes: str, default: str = "") -> str:
        for prefix in prefixes:
            prefix_norm = normalize_label(prefix)
            for key, val in norm_fields.items():
                if key.startswith(prefix_norm) and val:
                    return val
        return default

    url = get_field("URL")
    briefing = {
        "nicho": get_field("Nicho"),
        "publico_alvo": get_by_prefix("Publico alvo"),
        "audiencia_alvos": get_by_prefix("Exemplo de Audiencia"),
        "b2b_b2c": get_field("B2B ou B2C"),
        "escopo": get_by_prefix("Local ou Nacional"),
        "tipo_negocio": get_by_prefix("Servico, Digital, Ecommerce"),
        "modelo_produto": get_field("Modelo de Produto"),
        "dores": get_by_prefix("Dores que resolve"),
        "ferramenta_saas": get_field("Ferramenta/SaaS"),
        "curso": get_field("Curso"),
        "ai": get_field("AI"),
        "nicho_principal": get_by_prefix("Nicho principal"),
        "nicho_secundario": get_by_prefix("Nicho secundario"),
        "seo_keywords": get_by_prefix("Defina 2 palavras para SEO"),
        "nicho_amplo": get_by_prefix("Nicho Amplo"),
        "nicho_medio": get_by_prefix("Nicho Medio"),
        "nicho_especifico": get_by_prefix("Nicho Especifico"),
        "produto_principal": get_by_prefix("Produto principal"),
        "produtos": get_by_prefix("Produtos"),
        "perfil": get_by_prefix("Perfil"),
        "cidades": [c.strip() for c in re.split(r"[,;]", get_by_prefix("Cidades alvo")) if c.strip()],
        "termos_raiz": [t.strip() for t in re.split(r"[,;]", get_by_prefix("Termos raiz")) if t.strip()],
        "personas": [],
        "canais": [],
    }

    if "Personas & Canais" in wb.sheetnames:
        wsp = wb["Personas & Canais"]
        mode = ""
        for row in wsp.iter_rows(values_only=True):
            vals = [("" if v is None else str(v).strip()) for v in row]
            if not any(vals):
                continue
            first = vals[0]
            if first == "Personas":
                mode = "personas"
                continue
            if first == "Potencial de Canais de Marketing":
                mode = "canais"
                continue
            if mode == "personas" and first == "Persona":
                continue
            if mode == "canais" and first == "Canal":
                continue
            if mode == "personas":
                briefing["personas"].append({
                    "nome": vals[0],
                    "idade_perfil": vals[1] if len(vals) > 1 else "",
                    "descricao": vals[2] if len(vals) > 2 else "",
                    "onde_busca": vals[3] if len(vals) > 3 else "",
                })
            elif mode == "canais":
                briefing["canais"].append({
                    "canal": vals[0],
                    "conversao": vals[1] if len(vals) > 1 else "",
                    "velocidade": vals[2] if len(vals) > 2 else "",
                    "custo": vals[3] if len(vals) > 3 else "",
                    "justificativa": vals[4] if len(vals) > 4 else "",
                })

    traducao = empty_traducao()
    if "Traducao" in wb.sheetnames:
        wst = wb["Traducao"]
        tfields = {}
        for row in wst.iter_rows(values_only=True):
            if not row or row[0] is None:
                continue
            key = str(row[0]).strip()
            val = "" if len(row) < 2 or row[1] is None else str(row[1]).strip()
            tfields[key] = val
        tnorm = {normalize_label(k): v for k, v in tfields.items()}

        def tget(prefix: str) -> str:
            prefix_norm = normalize_label(prefix)
            for key, val in tnorm.items():
                if key.startswith(prefix_norm) and val:
                    return val
            return ""

        traducao["nicho"] = {
            "amplo": tget("Nicho Amplo") or "—",
            "medio": tget("Nicho Medio") or "—",
            "especifico": tget("Nicho Especifico") or "—",
        }
        traducao["produto"] = {
            "amplo": tget("Produto Amplo") or "—",
            "medio": tget("Produto Medio") or "—",
            "especifico": tget("Produto Especifico") or "—",
            "exemplo_dor": tget("Produto exemplo dor") or "—",
            "antes": tget("Produto antes") or "—",
            "depois": tget("Produto depois") or "—",
            "beneficios": tget("Produto beneficios") or "—",
        }
        traducao["publico_alvo"] = tget("Publico alvo") or "—"
        traducao["audiencia_empresas"] = [c.strip() for c in re.split(r"[,;]", tget("Audiencia Empresas")) if c.strip()]
        traducao["audiencia_cargos"] = [c.strip() for c in re.split(r"[,;]", tget("Audiencia Cargos")) if c.strip()]
        traducao["dores_exemplos"] = [c.strip() for c in re.split(r"\|", tget("Dores exemplos")) if c.strip()]
        traducao["dores_que_resolve"] = [c.strip() for c in re.split(r"\|", tget("Dores que resolve")) if c.strip()]
        traducao["motivacoes"] = [c.strip() for c in re.split(r"\|", tget("Motivacoes")) if c.strip()]
        traducao["onde_busca_informacao"] = [c.strip() for c in re.split(r"\|", tget("Onde busca informacao")) if c.strip()]

    if "Traducao Produtos" in wb.sheetnames:
        wsp2 = wb["Traducao Produtos"]
        for idx, row in enumerate(wsp2.iter_rows(values_only=True)):
            if idx == 0:
                continue
            vals = [("" if v is None else str(v).strip()) for v in row]
            if not any(vals):
                continue
            traducao["produtos"].append({
                "nome": vals[0] if len(vals) > 0 else "",
                "amplo": vals[1] if len(vals) > 1 else "",
                "medio": vals[2] if len(vals) > 2 else "",
                "especifico": vals[3] if len(vals) > 3 else "",
                "exemplo_dor": vals[4] if len(vals) > 4 else "",
                "antes": vals[5] if len(vals) > 5 else "",
                "depois": vals[6] if len(vals) > 6 else "",
                "beneficios": vals[7] if len(vals) > 7 else "",
            })

    briefing["traducao"] = traducao

    briefing["formula_canais"] = {}
    if "Formula Canais" in wb.sheetnames:
        wsf = wb["Formula Canais"]
        ffields = {}
        for row in wsf.iter_rows(values_only=True):
            if not row or row[0] is None:
                continue
            key = str(row[0]).strip()
            val = "" if len(row) < 2 or row[1] is None else str(row[1]).strip()
            ffields[key] = val
        briefing["formula_canais"] = {
            "produto": {
                "nicho": ffields.get("Produto / nicho", "").split("|")[-1].strip() if ffields.get("Produto / nicho") else "",
                "produto": ffields.get("Produto / nicho", "").split("|")[0].strip() if ffields.get("Produto / nicho") else "",
                "compra_real": ffields.get("Compra real", ""),
                "emocao_antes": ffields.get("Emocao antes", ""),
                "emocao_depois": ffields.get("Emocao depois", ""),
                "observacao": ffields.get("Observacao", ""),
            },
            "top_motores": [],
            "contexto": {},
            "ranking_potencial": [],
            "ranking_oportunidade": [],
            "estrategia": {
                "curto": [{"canal": x.strip()} for x in ffields.get("Curto prazo", "").split(",") if x.strip()],
                "medio": [{"canal": x.strip()} for x in ffields.get("Medio prazo", "").split(",") if x.strip()],
                "longo": [{"canal": x.strip()} for x in ffields.get("Longo prazo", "").split(",") if x.strip()],
            },
        }
        for part in ffields.get("Motores", "").split(","):
            if ":" in part:
                k, v = part.split(":", 1)
                try:
                    briefing["formula_canais"]["top_motores"].append((k.strip(), int(v.strip())))
                except Exception:
                    pass
        for part in ffields.get("Contexto", "").split(","):
            if "=" in part:
                k, v = part.split("=", 1)
                briefing["formula_canais"]["contexto"][k.strip()] = v.strip()
        for part in ffields.get("Top Potencial", "").split("|"):
            part = part.strip()
            m = re.match(r"(.+?) \(([-+0-9.]+)\)$", part)
            if m:
                briefing["formula_canais"]["ranking_potencial"].append({"canal": m.group(1), "classe": "", "potencial": m.group(2)})
        for part in ffields.get("Top Oportunidade", "").split("|"):
            part = part.strip()
            m = re.match(r"(.+?) \(([-+0-9.]+)\)$", part)
            if m:
                briefing["formula_canais"]["ranking_oportunidade"].append({"canal": m.group(1), "classe": "", "oportunidade": m.group(2)})

    wb.close()
    return url, briefing


def write_html(out_path: Path, url: str, briefing: dict):
    print(f"[OUT] Escrevendo HTML: {out_path}")
    domain = urlparse(url).netloc or url
    escopo = briefing.get("escopo", "")
    cidades = briefing.get("cidades") or []
    cidades_str = f" &middot; {', '.join(cidades)}" if cidades else ""
    personas = briefing.get("personas") or []
    canais = briefing.get("canais") or []

    html = HTML_TEMPLATE.format(
        domain=domain,
        url=url,
        nicho=briefing.get("nicho", "") or "—",
        nicho_principal=briefing.get("nicho_principal", "") or "—",
        escopo=escopo,
        cidades_str=cidades_str,
        modelo_produto=briefing.get("modelo_produto", "") or "—",
        b2b_b2c=briefing.get("b2b_b2c", "") or "—",
        tipo_negocio=briefing.get("tipo_negocio", "") or "—",
        n_personas=len(personas),
        n_canais=len(canais),
        date_str=datetime.now().strftime("%d/%m/%Y"),
        principal_section=build_principal_section_html(briefing),
        extra_section=build_extra_section_html(briefing),
        traducao_section=build_traducao_section_html(briefing),
        formula_section=build_formula_section_html(briefing),
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(html, encoding="utf-8")
    print(f"[OUT] HTML salvo: {out_path}")
    return out_path


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def normalize_url(url: str) -> str:
    url = url.strip()
    if not url:
        return ""
    if not url.startswith("http"):
        url = "https://" + url
    return url


def read_urls_from_file(path: Path) -> list:
    if not path.exists():
        print(f"[!] Arquivo de input nao encontrado: {path}")
        return []

    urls = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        url = normalize_url(line)
        if url:
            urls.append(url)
    return urls


def run(url: str, mvp: bool = True):
    t0 = time.time()
    slug = domain_slug(url)
    out_dir = OUTPUT_DIR / slug
    print(f"\n=== Pipeline Briefing Cliente: {url} ===")
    print(f"Output dir: {out_dir}")
    print(f"Modo: {'MVP (scrape + briefing)' if mvp else 'completo'}\n")

    site = scrape_site(url)
    if not site.get("body"):
        print("[!] Site vazio. Abortando.")
        return

    briefing = llm_briefing(url, site)
    print(f"[1B] Briefing: escopo={briefing['escopo']}, nicho='{briefing.get('nicho_principal')}', "
          f"cidades={briefing.get('cidades')}, termos_raiz={briefing.get('termos_raiz')}")

    if not mvp:
        briefing.update(llm_classificar2(url, site, briefing))
        print(f"[1B2] Classificar 2: amplo='{briefing.get('nicho_amplo')}', "
              f"medio='{briefing.get('nicho_medio')}', perfil='{briefing.get('perfil')}'")

        briefing["traducao"] = llm_traducao_homepage(url, site, briefing)
        print(f"[1D] Traducao: nicho='{briefing['traducao']['nicho'].get('especifico')}', "
              f"produto='{briefing['traducao']['produto'].get('especifico')}', "
              f"produtos={len(briefing['traducao'].get('produtos') or [])}")

        briefing["formula_canais"] = llm_formula_canais(url, briefing)
        print(f"[1E] Formula: top potencial='{(briefing['formula_canais'].get('ranking_potencial') or [{}])[0].get('canal', '—')}'")

        persona_canais = llm_personas_canais(url, site, briefing)
        briefing["personas"] = persona_canais.get("personas") or []
        briefing["canais"] = persona_canais.get("canais") or []
    else:
        # Campos vazios para o XLSX/HTML nao quebrarem
        briefing["traducao"] = empty_traducao()
        briefing.setdefault("formula_canais", {})
        briefing.setdefault("personas", [])
        briefing.setdefault("canais", [])

    xlsx_path = out_dir / f"briefing-cliente-{slug}.xlsx"
    html_path = out_dir / f"briefing-cliente-{slug}.html"
    write_xlsx(xlsx_path, url, briefing)
    write_html(html_path, url, briefing)

    elapsed = time.time() - t0
    print(f"\n=== Concluido em {elapsed:.1f}s ===")
    print(f"XLSX: {xlsx_path}")
    print(f"HTML: {html_path}")


def main():
    argv = [a for a in sys.argv[1:]]
    mvp = "--full" not in argv
    argv = [a for a in argv if a != "--full"]
    arg = argv[0].strip() if argv else ""

    if not arg:
        input_path = DEFAULT_INPUT_FILE
        urls = read_urls_from_file(input_path)
        if not urls:
            print("Uso:")
            print("  python \"1- entender o cliente.py\" https://empresa.com.br/")
            print("  python \"1- entender o cliente.py\" https://empresa.com.br/ --full")
            print(f"\nColoque as URLs em: {DEFAULT_INPUT_FILE}")
            sys.exit(1)
        print(f"[INPUT] {len(urls)} URL(s) lidas de {input_path}")
    elif arg.lower().endswith(".txt"):
        input_path = Path(arg)
        if not input_path.is_absolute():
            input_path = BASE_DIR / input_path
        urls = read_urls_from_file(input_path)
        if not urls:
            sys.exit(1)
        print(f"[INPUT] {len(urls)} URL(s) lidas de {input_path}")
    else:
        urls = [normalize_url(arg)]

    for i, url in enumerate(urls, 1):
        if len(urls) > 1:
            print(f"\n--- Cliente {i}/{len(urls)} ---")
        run(url, mvp=mvp)


if __name__ == "__main__":
    main()
