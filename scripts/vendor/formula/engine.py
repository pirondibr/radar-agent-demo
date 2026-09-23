"""
engine.py — Núcleo da metodologia "Fórmula de Potencial de Canais".

Implementa o Algoritmo Universal de Compatibilidade de Canais (Artigos 1 a 9):

    Empresa -> Produto -> Motores Psicológicos -> Variáveis Contextuais
    -> Motores Ajustados -> Compatibilidade de Canal -> Potencial Teórico
    -> Execução (Mercado Real) -> Oportunidade -> Estratégia

As pessoas não compram produtos. As pessoas compram motores psicológicos.
Os canais são ambientes que maximizam determinados motores.
As oportunidades surgem da diferença entre potencial e adoção do mercado.
"""

from __future__ import annotations

from dataclasses import dataclass, field


# ---------------------------------------------------------------------------
# 1. TAXONOMIA DE MOTORES PSICOLÓGICOS (Artigo 2)
# ---------------------------------------------------------------------------
# Cada motor pertence a uma das 8 classes. Usado apenas para exibição/organização.
MOTORES = {
    # Classe 1 — Sobrevivência
    "necessidade": "Sobrevivência",
    "urgencia": "Sobrevivência",
    "seguranca": "Sobrevivência",
    "economia": "Sobrevivência",
    "conveniencia": "Sobrevivência",
    # Classe 2 — Confiança
    "confianca": "Confiança",
    "autoridade": "Confiança",
    "prova_social": "Confiança",
    "proximidade": "Confiança",
    "reputacao": "Confiança",
    # Classe 3 — Crescimento
    "aspiracao": "Crescimento",
    "ganho_financeiro": "Crescimento",
    "escala": "Crescimento",
    "performance": "Crescimento",
    "liberdade": "Crescimento",
    # Classe 4 — Sociais
    "comunidade": "Social",
    "status": "Social",
    "reconhecimento": "Social",
    "influencia": "Social",
    "pertencimento": "Social",
    # Classe 5 — Identidade
    "identidade": "Identidade",
    "proposito": "Identidade",
    "significado": "Identidade",
    "expressao": "Identidade",
    "valores": "Identidade",
    # Classe 6 — Descoberta
    "curiosidade": "Descoberta",
    "descoberta": "Descoberta",
    "inovacao": "Descoberta",
    "surpresa": "Descoberta",
    "tendencia": "Descoberta",
    # Classe 7 — Transformação
    "transformacao": "Transformação",
    "autoestima": "Transformação",
    # Classe 8 — Cognitivos
    "educacao": "Cognitivo",
    "entretenimento": "Cognitivo",
    "networking": "Cognitivo",
}

# Rótulos amigáveis para exibição
MOTOR_LABEL = {
    "necessidade": "Necessidade",
    "urgencia": "Urgência",
    "seguranca": "Segurança",
    "economia": "Economia",
    "conveniencia": "Conveniência",
    "confianca": "Confiança",
    "autoridade": "Autoridade",
    "prova_social": "Prova Social",
    "proximidade": "Proximidade",
    "reputacao": "Reputação",
    "aspiracao": "Aspiração",
    "ganho_financeiro": "Ganho Financeiro",
    "escala": "Escala",
    "performance": "Performance",
    "liberdade": "Liberdade",
    "comunidade": "Comunidade",
    "status": "Status",
    "reconhecimento": "Reconhecimento",
    "influencia": "Influência",
    "pertencimento": "Pertencimento",
    "identidade": "Identidade",
    "proposito": "Propósito",
    "significado": "Significado",
    "expressao": "Expressão",
    "valores": "Valores",
    "curiosidade": "Curiosidade",
    "descoberta": "Descoberta",
    "inovacao": "Inovação",
    "surpresa": "Surpresa",
    "tendencia": "Tendência",
    "transformacao": "Transformação",
    "autoestima": "Autoestima",
    "educacao": "Educação",
    "entretenimento": "Entretenimento",
    "networking": "Networking",
}


# ---------------------------------------------------------------------------
# 2. TAXONOMIA DE CANAIS + MATRIZ CANAL × MOTOR (Artigos 3 e 6)
# ---------------------------------------------------------------------------
# Score de compatibilidade de cada canal com cada motor (0 = incompatível, 10 = máximo).
# "autoestima" é tratada como faceta de transformação (herda o score de transformacao).
# "horizonte" = papel natural do canal na estratégia (Artigo 9):
#   curto = caixa agora (mídia paga de resposta)
#   medio = vantagem (parcerias/descoberta que se constroem em semanas)
#   longo = monopólio (orgânico, autoridade, marca, SEO)
CANAIS: dict[str, dict] = {
    "Google (SEO)": {
        "classe": "Busca",
        "horizonte": "longo",
        "matriz": {
            "necessidade": 10, "urgencia": 8, "confianca": 8, "educacao": 9,
            "economia": 7, "autoridade": 8, "proximidade": 6, "seguranca": 7,
            "aspiracao": 2, "curiosidade": 1, "transformacao": 3, "comunidade": 1,
            "status": 1, "descoberta": 2, "reputacao": 7,
        },
    },
    "Google Ads": {
        "classe": "Busca",
        "horizonte": "curto",
        "matriz": {
            "necessidade": 10, "urgencia": 10, "economia": 8, "confianca": 8,
            "proximidade": 7, "seguranca": 7, "conveniencia": 7, "aspiracao": 2,
            "transformacao": 3, "curiosidade": 1, "descoberta": 2,
        },
    },
    "Google Maps": {
        "classe": "Busca",
        "horizonte": "curto",
        "matriz": {
            "proximidade": 10, "necessidade": 9, "confianca": 9, "urgencia": 8,
            "reputacao": 8, "prova_social": 8, "conveniencia": 7, "aspiracao": 1,
            "curiosidade": 0, "transformacao": 2,
        },
    },
    "Meta Ads": {
        "classe": "Aspiração/Pago",
        "horizonte": "curto",
        "matriz": {
            "aspiracao": 10, "economia": 9, "transformacao": 9, "descoberta": 8,
            "comunidade": 8, "confianca": 7, "curiosidade": 7, "status": 7,
            "ganho_financeiro": 8, "autoridade": 6, "necessidade": 5, "identidade": 6,
            "conveniencia": 6, "surpresa": 6,
        },
    },
    "Instagram Orgânico": {
        "classe": "Aspiração",
        "horizonte": "longo",
        "matriz": {
            "transformacao": 10, "aspiracao": 10, "status": 9, "identidade": 9,
            "comunidade": 8, "expressao": 8, "tendencia": 7, "confianca": 6,
            "reconhecimento": 7, "necessidade": 2, "curiosidade": 5,
        },
    },
    "TikTok": {
        "classe": "Descoberta",
        "horizonte": "medio",
        "matriz": {
            "descoberta": 10, "curiosidade": 10, "tendencia": 9, "entretenimento": 9,
            "inovacao": 9, "surpresa": 9, "transformacao": 8, "aspiracao": 6,
            "confianca": 3, "necessidade": 1,
        },
    },
    "YouTube": {
        "classe": "Educação",
        "horizonte": "longo",
        "matriz": {
            "educacao": 10, "autoridade": 10, "confianca": 9, "significado": 9,
            "aspiracao": 7, "comunidade": 6, "transformacao": 6, "performance": 7,
            "reputacao": 7, "necessidade": 3,
        },
    },
    "Influenciadores": {
        "classe": "Confiança",
        "horizonte": "medio",
        "matriz": {
            "confianca": 10, "prova_social": 9, "aspiracao": 9, "status": 8,
            "identidade": 8, "comunidade": 8, "transformacao": 8, "descoberta": 7,
            "tendencia": 7, "reputacao": 7,
        },
    },
    "LinkedIn": {
        "classe": "Autoridade",
        "horizonte": "medio",
        "matriz": {
            "autoridade": 10, "status": 9, "networking": 8, "educacao": 8,
            "escala": 8, "ganho_financeiro": 8, "economia": 7, "performance": 7,
            "reputacao": 7, "transformacao": 2,
        },
    },
    "WhatsApp / Indicações": {
        "classe": "Conversão/Confiança",
        "horizonte": "medio",
        "matriz": {
            "confianca": 10, "proximidade": 9, "conveniencia": 9, "prova_social": 9,
            "comunidade": 7, "urgencia": 7, "seguranca": 6,
        },
    },
    "Podcasts": {
        "classe": "Educação",
        "horizonte": "longo",
        "matriz": {
            "educacao": 10, "autoridade": 9, "significado": 9, "comunidade": 8,
            "confianca": 8, "proposito": 8, "aspiracao": 6,
        },
    },
}


# ---------------------------------------------------------------------------
# 3. DADOS DE EXECUÇÃO (Artigo 4) — escala 1 a 10
# ---------------------------------------------------------------------------
# dificuldade  : quão difícil é executar bem (maior = pior)
# velocidade   : tempo para retorno (maior = mais rápido)
# previsibilidade: consistência do resultado (maior = melhor)
# competicao   : intensidade competitiva padrão (maior = pior)
EXECUCAO: dict[str, dict[str, int]] = {
    "Google (SEO)":          {"dificuldade": 8, "velocidade": 3, "previsibilidade": 5, "competicao": 6},
    "Google Ads":            {"dificuldade": 6, "velocidade": 10, "previsibilidade": 10, "competicao": 8},
    "Google Maps":           {"dificuldade": 4, "velocidade": 6, "previsibilidade": 7, "competicao": 5},
    "Meta Ads":              {"dificuldade": 5, "velocidade": 10, "previsibilidade": 9, "competicao": 5},
    "Instagram Orgânico":    {"dificuldade": 9, "velocidade": 1, "previsibilidade": 4, "competicao": 8},
    "TikTok":                {"dificuldade": 10, "velocidade": 4, "previsibilidade": 2, "competicao": 7},
    "YouTube":               {"dificuldade": 9, "velocidade": 1, "previsibilidade": 5, "competicao": 6},
    "Influenciadores":       {"dificuldade": 3, "velocidade": 7, "previsibilidade": 6, "competicao": 5},
    "LinkedIn":              {"dificuldade": 6, "velocidade": 5, "previsibilidade": 6, "competicao": 4},
    "WhatsApp / Indicações": {"dificuldade": 1, "velocidade": 8, "previsibilidade": 7, "competicao": 3},
    "Podcasts":              {"dificuldade": 8, "velocidade": 2, "previsibilidade": 4, "competicao": 4},
}


# ---------------------------------------------------------------------------
# 4. VARIÁVEIS CONTEXTUAIS (Artigo 7)
# ---------------------------------------------------------------------------
# Cada opção de contexto produz multiplicadores por canal (centrados em 1.0).
# O que não aparece fica em 1.0 (neutro).
CONTEXTO_MULT: dict[str, dict[str, dict[str, float]]] = {
    "idade": {
        "jovem":  {"TikTok": 1.35, "Instagram Orgânico": 1.20, "YouTube": 1.15, "Google (SEO)": 0.85, "Google Ads": 0.85},
        "adulto": {"Google (SEO)": 1.15, "Google Ads": 1.15, "Meta Ads": 1.15, "Instagram Orgânico": 1.10, "YouTube": 1.05},
        "maduro": {"Meta Ads": 1.30, "Google (SEO)": 1.15, "Google Ads": 1.15, "YouTube": 1.05, "TikTok": 0.70},
        "senior": {"Meta Ads": 1.30, "Google Ads": 1.15, "WhatsApp / Indicações": 1.25, "TikTok": 0.45, "Instagram Orgânico": 0.75},
    },
    "modelo": {
        "B2C": {"TikTok": 1.20, "Instagram Orgânico": 1.20, "Meta Ads": 1.15, "LinkedIn": 0.40},
        "B2B": {"Google (SEO)": 1.20, "Google Ads": 1.15, "LinkedIn": 1.60, "Podcasts": 1.20, "TikTok": 0.45, "Instagram Orgânico": 0.75},
    },
    "alcance": {
        "local":    {"Google Maps": 1.60, "Google (SEO)": 1.20, "Google Ads": 1.15, "WhatsApp / Indicações": 1.25, "Meta Ads": 1.05, "TikTok": 0.80},
        "regional": {"Google Maps": 1.20, "Meta Ads": 1.10, "Google Ads": 1.10, "Instagram Orgânico": 1.05},
        "nacional": {"Google Maps": 0.20, "Meta Ads": 1.20, "TikTok": 1.25, "Instagram Orgânico": 1.15, "YouTube": 1.10},
    },
    "ticket": {
        "baixo": {"TikTok": 1.20, "Meta Ads": 1.15, "Instagram Orgânico": 1.05, "YouTube": 0.85},
        "medio": {"Meta Ads": 1.10, "Google Ads": 1.05, "Instagram Orgânico": 1.05},
        "alto":  {"Google (SEO)": 1.10, "YouTube": 1.20, "Influenciadores": 1.15, "LinkedIn": 1.05, "WhatsApp / Indicações": 1.10, "TikTok": 0.85},
    },
    "visualidade": {
        "baixa": {"Google (SEO)": 1.20, "Google Ads": 1.15, "LinkedIn": 1.10, "Instagram Orgânico": 0.75, "TikTok": 0.65},
        "media": {"Meta Ads": 1.05, "Instagram Orgânico": 1.05},
        "alta":  {"Instagram Orgânico": 1.30, "TikTok": 1.30, "Meta Ads": 1.15, "Influenciadores": 1.15, "Google (SEO)": 0.90},
    },
    "confianca_risco": {  # quanto custa errar
        "baixo":      {},
        "medio":      {"Google (SEO)": 1.05, "Influenciadores": 1.05},
        "alto":       {"Google (SEO)": 1.15, "YouTube": 1.15, "Influenciadores": 1.10, "WhatsApp / Indicações": 1.15, "TikTok": 0.85},
        "muito_alto": {"Google (SEO)": 1.20, "YouTube": 1.20, "WhatsApp / Indicações": 1.25, "Google Maps": 1.10, "TikTok": 0.70},
    },
    "surpresa": {
        "baixa": {"Google (SEO)": 1.10, "Google Ads": 1.05, "TikTok": 0.80},
        "media": {"Meta Ads": 1.05, "Instagram Orgânico": 1.05},
        "alta":  {"TikTok": 1.30, "Meta Ads": 1.15, "Influenciadores": 1.15, "Instagram Orgânico": 1.10, "Google (SEO)": 0.80},
    },
    "dor_desejo": {  # dor = resolver problema | desejo = tornar-se algo
        "dor":    {"Google (SEO)": 1.20, "Google Ads": 1.20, "Google Maps": 1.15, "WhatsApp / Indicações": 1.05, "TikTok": 0.75},
        "misto":  {"Meta Ads": 1.05},
        "desejo": {"Instagram Orgânico": 1.25, "TikTok": 1.20, "Meta Ads": 1.15, "Influenciadores": 1.20, "Google (SEO)": 0.85},
    },
    "potencial_organico": {  # o nicho produz conteúdo naturalmente?
        "baixo": {"Meta Ads": 1.10, "Google Ads": 1.10, "Instagram Orgânico": 0.85, "TikTok": 0.85},
        "medio": {},
        "alto":  {"Instagram Orgânico": 1.20, "TikTok": 1.20, "YouTube": 1.15, "Influenciadores": 1.10},
    },
}


@dataclass
class Produto:
    """Perfil de entrada: nicho > empresa > produto + motores + contexto."""
    nicho: str
    empresa: str
    produto: str
    # o que a pessoa realmente compra / emoção antes / emoção depois (Artigo 5)
    compra_real: str
    emocao_antes: str
    emocao_depois: str
    # motores psicológicos com peso 0-10 (Artigo 5/6)
    motores: dict[str, int]
    # variáveis contextuais (Artigo 7)
    contexto: dict[str, str]
    observacao: str = ""


# ---------------------------------------------------------------------------
# 5. MOTOR DE CÁLCULO (Artigo 9 — etapas 4 a 8)
# ---------------------------------------------------------------------------
def _mult_contexto(canal: str, contexto: dict[str, str]) -> float:
    """Multiplicador contextual combinado de um canal."""
    mult = 1.0
    for variavel, valor in contexto.items():
        tabela = CONTEXTO_MULT.get(variavel, {})
        mult *= tabela.get(valor, {}).get(canal, 1.0)
    return mult


def _compat_base(canal: str, motores: dict[str, int]) -> float:
    """Compatibilidade psicológica pura do canal com o perfil de motores (0-100)."""
    matriz = CANAIS[canal]["matriz"]
    raw = 0.0
    max_raw = 0.0
    for motor, peso in motores.items():
        # autoestima herda score de transformação quando o canal não a lista
        score = matriz.get(motor)
        if score is None and motor == "autoestima":
            score = matriz.get("transformacao", 0)
        score = score if score is not None else 0
        raw += peso * score
        max_raw += peso * 10
    if max_raw == 0:
        return 0.0
    return raw / max_raw * 100.0


def analisar(produto: Produto) -> dict:
    """Executa o algoritmo completo para um produto e devolve o resultado estruturado."""
    canais_result = []

    # Etapa 4/5 — compatibilidade + potencial teórico (com contexto)
    for canal in CANAIS:
        base = _compat_base(canal, produto.motores)
        mult = _mult_contexto(canal, produto.contexto)
        potencial_raw = base * mult  # pode passar de 100; normalizado abaixo

        exe = EXECUCAO[canal]
        # Fórmula do Canal Real (Artigo 4):
        # Potencial × Velocidade × Previsibilidade ÷ Dificuldade ÷ Competição.
        # Cada fator é comprimido para uma faixa limitada, para que a execução
        # MODULE o potencial (em vez de um único fator dominar por ordens de
        # grandeza, ex.: dificuldade=1 valendo 10x).
        vel_f = 0.60 + 0.40 * (exe["velocidade"] / 10.0)       # 0.64 .. 1.00
        prev_f = 0.60 + 0.40 * (exe["previsibilidade"] / 10.0)  # 0.64 .. 1.00
        dif_f = 1.30 - 0.60 * (exe["dificuldade"] / 10.0)       # 0.70 .. 1.24
        comp_f = 1.20 - 0.50 * (exe["competicao"] / 10.0)       # 0.70 .. 1.15
        fator_exec = vel_f * prev_f * dif_f * comp_f

        canais_result.append({
            "canal": canal,
            "classe": CANAIS[canal]["classe"],
            "horizonte": CANAIS[canal]["horizonte"],
            "base": round(base, 1),
            "mult_contexto": round(mult, 2),
            "potencial_raw": potencial_raw,
            "fator_exec": fator_exec,
            "execucao": exe,
        })

    # Normalização relativa dentro do produto: o melhor canal = 100.
    # Isso preserva a diferenciação entre canais (evita empates no teto de 100).
    max_pot_raw = max(c["potencial_raw"] for c in canais_result) or 1.0
    for c in canais_result:
        c["potencial"] = round(c["potencial_raw"] / max_pot_raw * 100.0, 1)
        c["mercado_raw"] = c["potencial"] * c["fator_exec"]

    # Etapa 7 — "Mercado Real" também normalizado (o mais usado = 100)
    max_merc_raw = max(c["mercado_raw"] for c in canais_result) or 1.0
    for c in canais_result:
        c["mercado"] = round(c["mercado_raw"] / max_merc_raw * 100.0, 1)
        # Etapa 8 — Oportunidade = Potencial - Adoção do Mercado
        c["oportunidade"] = round(c["potencial"] - c["mercado"], 1)

    ranking_potencial = sorted(canais_result, key=lambda c: c["potencial"], reverse=True)
    ranking_mercado = sorted(canais_result, key=lambda c: c["mercado"], reverse=True)
    ranking_oportunidade = sorted(canais_result, key=lambda c: c["oportunidade"], reverse=True)

    estrategia = _montar_estrategia(canais_result)

    top_motores = sorted(produto.motores.items(), key=lambda kv: kv[1], reverse=True)

    return {
        "produto": produto,
        "top_motores": top_motores,
        "canais": canais_result,
        "ranking_potencial": ranking_potencial,
        "ranking_mercado": ranking_mercado,
        "ranking_oportunidade": ranking_oportunidade,
        "estrategia": estrategia,
    }


def _montar_estrategia(canais: list[dict]) -> dict[str, list[dict]]:
    """Etapa 9 — agrupa os canais relevantes pelo seu horizonte natural (Artigo 9).

    Curto prazo  : gera caixa agora  -> mídia paga de resposta (Meta/Google Ads/Maps).
    Médio prazo  : gera vantagem     -> parcerias/descoberta (Influenciadores, TikTok...).
    Longo prazo  : gera monopólio    -> orgânico, autoridade e marca (SEO, YouTube, IG).

    Só entram canais com potencial relevante; ordenados pela métrica que importa
    em cada horizonte (mercado no curto, oportunidade no médio/longo).
    """
    GATE = 25.0  # potencial mínimo para um canal entrar na estratégia
    buckets: dict[str, list[dict]] = {"curto": [], "medio": [], "longo": []}
    for c in canais:
        if c["potencial"] >= GATE:
            buckets[c["horizonte"]].append(c)

    buckets["curto"].sort(key=lambda c: c["mercado"], reverse=True)
    buckets["medio"].sort(key=lambda c: c["oportunidade"], reverse=True)
    buckets["longo"].sort(key=lambda c: c["oportunidade"], reverse=True)
    return {k: v[:4] for k, v in buckets.items()}
