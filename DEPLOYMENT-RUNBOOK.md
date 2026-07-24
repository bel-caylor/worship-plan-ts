# Deployment Runbook

This is the step-by-step process for deploying this project when we have:

- local code changes in this repo
- an Apps Script web app deployment
- the Cloudflare Worker proxy in `worship-plan-proxy`
- the standalone frontend that should point at the Worker

## The N-1 / "N1" issue

This project has one easy-to-miss deployment trap:

- Apps Script deployment `N` gets a new deployment URL.
- The Worker secret still points at deployment `N-1` (the previous URL).
- The standalone app starts failing with `Invalid RPC response` or a non-JSON error.

In plain English: the Worker is live, but it is forwarding to the old Apps Script deployment.

Typical symptoms:

- `Invalid RPC response`
- `Apps Script returned a non-JSON response`
- the Worker returns HTML instead of JSON

Usually this means one of two things:

1. `APPS_SCRIPT_BASE` on the Worker still points to the old Apps Script deployment
2. the Apps Script web app was not deployed with `Who has access = Anyone`

## What points to what

Keep these straight:

- Apps Script Web App URL: this is the Google deployment URL
- Worker secret `APPS_SCRIPT_BASE`: this must point to the current Apps Script Web App URL
- Standalone frontend `APPS_SCRIPT_BASE`: this must point to the Worker URL, not the Apps Script URL

That distinction is the main thing to remember.

## Full deployment process

### 1. Deploy the Apps Script code from local

From the repo root:

```powershell
npm run deploy
```

What this does:

- rebuilds `dist/`
- pushes the latest Apps Script bundle with `clasp`

If you want the non-force version instead:

```powershell
npm run deploy:backend
```

## 2. Finish the deployment in the Apps Script UI

After the local push:

1. Open the Apps Script project
2. Go to `Deploy`
3. Choose `Manage deployments`
4. Create a new `Web app` deployment, or edit/update the current one
5. Set `Execute as` to `Me`
6. Set `Who has access` to `Anyone`
7. Save the deployment
8. Copy the deployment base URL

Use the base form:

```text
https://script.google.com/macros/s/<DEPLOYMENT_ID>
```

Notes:

- If Google shows you an `/exec` URL, that is also acceptable
- the Worker normalizes either form
- every time Apps Script gives you a new deployment URL, assume the Worker secret now needs updating too

## 3. Update the Wrangler secret

Change into the Worker folder:

```powershell
cd worship-plan-proxy
```

Set the secret:

```powershell
npx wrangler secret put APPS_SCRIPT_BASE
```

When prompted, paste the current Apps Script deployment URL from step 2.

Use:

```text
https://script.google.com/macros/s/<DEPLOYMENT_ID>
```

or:

```text
https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec
```

Both are fine for this Worker.

## 4. Redeploy the Worker

Still inside `worship-plan-proxy`:

```powershell
npx wrangler deploy
```

At this point:

- the Worker is live
- the Worker now points at the current Apps Script deployment

This is the step that fixes the N-1 issue.

## 5. Rebuild the standalone frontend against the Worker URL

Go back to the repo root and build the standalone site using the Worker URL, not the Apps Script URL.

PowerShell:

```powershell
$env:APPS_SCRIPT_BASE = "https://worship-plan-proxy.belinda-caylor.workers.dev"
npm run build:standalone
```

Important:

- for the standalone site, `APPS_SCRIPT_BASE` should be the Worker URL
- do not set the standalone build to the Apps Script Google URL

## Quick version

If you just want the shortest safe checklist:

1. From repo root: `npm run deploy`
2. In Apps Script UI: deploy/update Web app
3. In Apps Script UI: confirm `Execute as = Me`
4. In Apps Script UI: confirm `Who has access = Anyone`
5. Copy the new Apps Script deployment URL
6. In `worship-plan-proxy`: `npx wrangler secret put APPS_SCRIPT_BASE`
7. In `worship-plan-proxy`: `npx wrangler deploy`
8. From repo root: build standalone using the Worker URL

## Copy/paste command block

Repo root:

```powershell
npm run deploy
```

Worker folder:

```powershell
cd worship-plan-proxy
npx wrangler secret put APPS_SCRIPT_BASE
npx wrangler deploy
```

Repo root again:

```powershell
$env:APPS_SCRIPT_BASE = "https://worship-plan-proxy.belinda-caylor.workers.dev"
npm run build:standalone
```

## Troubleshooting

If the standalone app breaks after a fresh Apps Script deploy, check these in order:

1. Did Apps Script create or update the web app deployment successfully?
2. Did you copy the new Apps Script deployment URL?
3. Did you run `npx wrangler secret put APPS_SCRIPT_BASE` with that new URL?
4. Did you run `npx wrangler deploy` afterward?
5. Is the Apps Script deployment access set to `Anyone`?
6. Is the standalone build using the Worker URL instead of the Apps Script URL?

## Rule of thumb

Whenever the Apps Script deployment URL changes, always do both:

1. update `APPS_SCRIPT_BASE` in Wrangler
2. redeploy the Worker
