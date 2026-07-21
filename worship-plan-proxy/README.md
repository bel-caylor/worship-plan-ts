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
set, it falls back to the default URL in `src/index.ts`. The Worker normalizes
the value, so either of these work:

- `https://script.google.com/macros/s/<DEPLOYMENT_ID>`
- `https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec`

## Frontend configuration

Build the standalone UI with the Worker URL as the base:

```powershell
$env:APPS_SCRIPT_BASE = "https://worship-plan-proxy.belinda-caylor.workers.dev"
npm run build:standalone
```

The GitHub Pages workflow also reads the `APPS_SCRIPT_BASE` secret; update it
to the Worker URL, not the Apps Script URL, whenever the Worker hostname
changes.

## Redeploy checklist

When the Apps Script Web app is redeployed, its deployment URL can change. If
the Worker still points at the old deployment, the standalone app will usually
show `Invalid RPC response` or a non-JSON RPC error.

1. In Apps Script, update or create the Web app deployment.
2. Set `Execute as` to `Me`.
3. Set `Who has access` to `Anyone`.
4. Copy the new Apps Script deployment URL:
   `https://script.google.com/macros/s/<DEPLOYMENT_ID>` or `/exec`
5. In this folder, update the Wrangler secret:

```powershell
npx wrangler secret put APPS_SCRIPT_BASE
```

6. Redeploy the Worker:

```powershell
npx wrangler deploy
```

Important:
- The standalone frontend should point to the Worker URL.
- The Worker secret `APPS_SCRIPT_BASE` should point to the current Apps Script deployment URL.
- Missing the `Anyone` access setting on the Apps Script deployment will cause the Worker to receive an HTML Google page instead of JSON.
