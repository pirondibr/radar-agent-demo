# Radar Agent

Chatbot de análise competitiva (Briefing, Concorrentes, Google Ads, SEO, Marca).

## Live no Render (recomendado)

O app roda no **Render** (não Vercel): Flask + SSE + pipeline de ~10 min.

1. Deploy/push deste repo (auto-deploy se já conectado)
2. No serviço **radar-agent-demo** → **Environment**, defina:

| Key | Obrigatório |
|-----|-------------|
| `DEMO_ONLY` | `0` |
| `OPENROUTER_API_KEY` | sim (live) |
| `SCRAPINGBEE_API_KEY` | sim (live / extras) |
| `DATAFORSEO_USER` | sim (live) |
| `DATAFORSEO_PASS` | sim (live) |
| `SEMRUSH_API_KEY` | sim (live) |
| `MERCADOPAGO_ACCESS_TOKEN` | sim (pago Pro) |
| `MERCADOPAGO_PUBLIC_KEY` | recomendado |
| `PUBLIC_BASE_URL` | `https://seu-app.onrender.com` |

3. Save → aguarde o redeploy
4. Confira `/api/hello`: `live_ready` e `payments_ready`

### Pagamento Pro (Mercado Pago)

Checkout Pro com **PIX + cartão**. CTA "Desbloquear R$ 99" abre o checkout; após `approved`, captura e-mail/WhatsApp.

Webhook: `POST /api/webhooks/mercadopago`  
Return: `/pay/return`

MCP Cursor (credenciais/docs): em Settings → Tools & MCPs, conecte `mercadopago-mcp-server` (`https://mcp.mercadopago.com/mcp`) e autorize o país BR.

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
