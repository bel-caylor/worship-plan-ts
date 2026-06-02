# Worship Plan Proxy

This Cloudflare Worker forwards RPC calls from the standalone frontend to the
Apps Script backend. It lets the public site talk to Apps Script without
running inside the Apps Script iframe or fighting CORS.

## How it works

1. The browser `fetch`es the Worker URL (e.g.,
   `https://worship-plan-proxy.belinda-caylor.workers.dev`) with a body like
   `{ "method": "getSongsForView", "payload": null }`.
2. The Worker posts the same payload to the Apps Script web app
   (`https://script.google.com/macros/s/<ID>/exec`).
3. Apps Script returns `{ ok: true, data: … }`; the Worker copies that JSON
   back to the browser and adds the necessary CORS headers.

Admins who run the UI inside Apps Script continue to use `google.script.run`
and bypass the Worker.

## Local development

```bash
cd worship-plan-proxy
npm install          # once
wrangler login       # once per machine

# set the Apps Script deployment URL for the Worker:
wrangler secret put APPS_SCRIPT_BASE

# then deploy:
wrangler deploy
```

The Worker reads `APPS_SCRIPT_BASE` from a Wrangler secret. If no secret is
set, it falls back to the default URL in `src/index.ts`.

## Frontend configuration

Build the standalone UI with the Worker URL as the base:

```powershell
$env:APPS_SCRIPT_BASE = "https://worship-plan-proxy.belinda-caylor.workers.dev"
npm run build:standalone
```

The GitHub Pages workflow also reads the `APPS_SCRIPT_BASE` secret; update it
to the Worker URL, not the Apps Script URL, whenever the Worker hostname
changes.
