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

# edit src/index.ts then:
wrangler deploy
```

Set `APPS_SCRIPT_BASE` at the top of `src/index.ts` (or store it as a Worker
secret) so the proxy knows which Apps Script deployment to call.

## Frontend configuration

Build the standalone UI with the Worker URL as the base:

```powershell
$env:APPS_SCRIPT_BASE = "https://worship-plan-proxy.belinda-caylor.workers.dev"
npm run build:standalone
```

The GitHub Pages workflow also reads the `APPS_SCRIPT_BASE` secret; update it
whenever the Worker hostname changes.
