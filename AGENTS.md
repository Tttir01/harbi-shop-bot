# Working notes for this repo (Base44 sandbox)

## What the app is

A single-file Flask `app.py` that serves Telegram bot webhooks (`GET /`, `GET /health`,
`POST /telegram-webhook`). There is no HTML frontend — the preview shows JSON, which is
the app's real output. Long-running state (`PRODUCTS`, the parsed Excel catalogue) lives in
process memory, so restarting the container drops an uploaded workbook; that is by design
(the deploy target, Render, has no durable disk either).

## Running it here

```sh
docker compose -f docker-compose.base44.yml up -d --build
docker compose -f docker-compose.base44.yml logs -f web
```

- `web` is `python:3.12-slim` with the repo bind-mounted at `/app`; its startup command
  installs `requirements.txt` and then runs gunicorn with `--reload` on port 3000, so file
  edits are picked up without a rebuild (no manual restart needed).
- The first run of a fresh container spends ~20s installing pip dependencies; the healthcheck
  has a 120s `start_period` to cover that plus reload cycles.
- `/health` is the health endpoint used by both Render and the compose healthcheck. It is
  read-only and requires nothing but the app itself.

## Credentials

- `BOT_TOKEN` (from @BotFather) and optional `SERPAPI_API_KEY` come from the user via the
  Base44 dashboard → `/run/base44/app.env`. Placeholders for all keys live in the committed
  `.env.base44-defaults`, which is listed *before* the platform file in `env_file:` so a
  dashboard value always wins.
- `RENDER_EXTERNAL_URL` is only set on Render. While it is empty the app skips
  `setWebhook` at import time and logs a warning — deliberate, so a sandbox run never
  repoints the user's live Telegram webhook away from their Render service. To test the bot
  from Telegram against this sandbox, set `RENDER_EXTERNAL_URL=https://3000-$BASE44_PUBLIC_HOST_SUFFIX`
  in the sandbox; the app then registers the sandbox URL as the bot's webhook.

## Verifying without Telegram credentials

A fabricated Telegram update exercises the real handler path and touches no live chat
(a fake `chat_id` makes Telegram reject the outbound `sendMessage`, so nobody is messaged):

```sh
docker compose -f docker-compose.base44.yml exec -T web python - <<'PY'
import json, urllib.request, os
req = urllib.request.Request(
    "http://127.0.0.1:3000/telegram-webhook",
    data=json.dumps({"message": {"chat": {"id": 123456789}, "text": "/start"}}).encode(),
    headers={"Content-Type": "application/json", "X-Telegram-Bot-Api-Secret-Token": os.environ.get("WEBHOOK_SECRET", "")},
)
print(urllib.request.urlopen(req, timeout=10).read())
PY
```

A wrong/missing secret must return 403. `GET /` and `GET /health` must return JSON with
`status: ok`/`healthy`.

## Tests and linting

The repo ships no test suite, linter config or CI. There is nothing to run beyond booting the
app and curling the three routes.
