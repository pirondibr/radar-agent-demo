# Radar Agent

Chatbot de análise competitiva (Briefing, Concorrentes, Google Ads, SEO, Marca).

## Live no Render (recomendado)

O app roda no **Render** (não Vercel): Flask + SSE + pipeline de ~10 min.

1. Deploy/push deste repo (auto-deploy se já conectado)
2. No serviço **radar-agent-demo** → **Environment**, defina:

| Key | Obrigatório |
|-----|-------------|
| `DEMO_ONLY` | `0` |
| `OPENROUTER_API_KEY` | sim |
| `SCRAPINGBEE_API_KEY` | sim |
| `DATAFORSEO_USER` | sim |
| `DATAFORSEO_PASS` | sim |
| `SEMRUSH_API_KEY` | sim |

3. Salve → aguarde o redeploy
4. Abra a URL e teste com uma URL real (ex: `https://www.mendesortega.com.br/`)
5. Digite `demo` se quiser só o exemplo Chatguru (sem scrape)

Confira `/api/hello`: `live_ready: true` e `version: 1.1.0`.

## Rodar local

```bash
# Copie .env.example → .env e preencha as keys
pip install -r chatbot/requirements.txt
python chatbot/server.py
```

Abra http://127.0.0.1:8766

## Por que não Vercel?

Serverless com timeout curto. A pipeline live precisa de processo longo + SSE — use Render.

## Estrutura

```
chatbot/                 # Flask + UI
scripts/                 # runners 1 / 3 / 5a–c
scripts/vendor/          # seo_pipeline + find_concorrentes* (Linux-ready)
outputs/metricas/chatguru/  # seed do exemplo demo
```
