# Radar Agent — Demo

Chatbot de análise competitiva (Briefing, Concorrentes, Google Ads, SEO, Marca).

## Demo pública

Em produção o app sobe com `DEMO_ONLY=1`: qualquer mensagem roda o exemplo **Chatguru** (sem scrapes ao vivo). Ideal para validar UX antes da versão paga.

## Rodar local

```bash
cd chatbot
pip install -r requirements.txt
python server.py
```

Abra http://127.0.0.1:8766 — confira a versão no topo (`v1.0.7+`).

Pipeline ao vivo (Windows, com scripts + Spy):

```bash
# sem DEMO_ONLY
set OPENROUTER_API_KEY=...
set SCRAPINGBEE_API_KEY=...
python server.py
```

## Deploy (Render)

**Por que não Vercel?** Este app usa Flask + SSE + jobs de minutos. A Vercel é serverless com timeout curto — não serve para o pipeline. O demo público fica no **Render**.

1. Suba este repositório no GitHub
2. Em [Render](https://render.com) → New → Blueprint → selecione o repo (`render.yaml`)
3. Aguarde o deploy e use a URL `https://….onrender.com`

Ou: New → Web Service → conecte o repo, com:

- Build: `pip install -r chatbot/requirements.txt`
- Start: `gunicorn -b 0.0.0.0:$PORT --workers 1 --threads 8 --timeout 120 --chdir chatbot server:app`
- Env: `DEMO_ONLY=1`

## Estrutura

```
chatbot/          # Flask + UI
scripts/          # runners 1 / 3 / 5a–c (live local)
outputs/metricas/chatguru/  # seed do demo
```
