Worship Planner (Apps Script + Alpine)

Overview
- Spreadsheet‑backed worship planning tool built on Google Apps Script.
- Client UI rendered via HtmlService using Alpine.js.
- Clean separation between views (markup), apps (Alpine logic), and context (client constants).

Repo layout
- src/html/index.html        – root HTML template; loads head scripts and mounts routed views
- src/html/views/*.html      – view partials; markup/bindings only (no scripts)
- src/html/js/util.html      – client helpers (e.g., callRpc, notify)
- src/html/js/apps-*.html    – Alpine apps (songsApp, weeklyPlanApp)
- src/html/context.html      – client‑only constants (APP_CTX)
- src/*.ts                   – Apps Script server code (compiled to dist/Code.gs)
- src/constants.ts           – server constants (sheets, columns)
- src/features/*.ts          – server features (songs, leaders, services)
- dist/*                     – build output pushed by clasp

Getting started
1) Install deps: `npm i`
2) Dev loop: `npm run watch` (build + copy + push on change)
   - or one‑off: `npm run build && npx clasp push`
3) Deploy web app in Apps Script: Deploy → Manage deployments → New deployment → Web app

Authorizations
- Manifest is pushed as `dist/appsscript.json` and sets:
  - runtime V8
  - scopes for Sheets, Drive (readonly), external requests (for ESV)
- First time you use ESV preview/fetch, set `ESV_API_TOKEN` in Script Properties.

Context pack (client config)
- `src/html/context.html` defines `window.APP_CTX` with:
  - `defaults`: leader, sermon, serviceType, time
  - `leaderChoices`, `sermonChoices` (seed lists merged with dynamic ones from Sheets)
  - `bibleBooks`: `[ [Book, chapterCount], ... ]`
- Alpine apps read from APP_CTX during `init()`; views remain script‑free.

RPC pattern
- Client calls server via a single entrypoint: `callRpc(method, payload)` (util.html).
- Server routes in `src/rpc.ts` using a switch on `method`.
- Add a feature by implementing a function (e.g., `addService`) and wiring a case in `rpc.ts`.

Services sheet integration
- Server constants for columns: `SERVICES_COL` in `src/constants.ts`.
- Adding a service (`src/features/services.ts:addService`):
  - Generates `ServiceID` in format `YYYY-MM-DD_h[[:mm]]am|pm`.
  - Writes `Scripture` reference and, if a `Scripture Text` column exists, fetches passage text via ESV and stores cleaned content.
- People lists (`getServicePeople`) are normalized (trim/case/space collapse, title‑cased for display) to avoid duplicates like `Tom` vs ` tom`.

Development conventions
- Keep view files free of `<script>` and closing `</body>` tags.
- Add client constants in `context.html`, not inside app files.
- Keep server secrets (tokens) only in Script Properties; never in code or context.
- When adding an RPC, keep the `rpc.ts` switch as the single router.

Build scripts
- `npm run build` – bundle server, copy HTML, build CSS, wrap CSS
- `npm run watch` – builds continuously and pushes with clasp
- `npm run build:standalone` – full build + produce `dist-standalone/index.html` for static hosting
- `npx clasp push -f` – force push if necessary (e.g., after manifest change)

Standalone frontend option
1. Deploy the Apps Script Web App (Execute as you, access anyone with link) and note the base URL (`https://script.google.com/macros/s/<DEPLOYMENT_ID>`).
2. Build static assets: `APPS_SCRIPT_BASE=<base url> npm run build:standalone`. This writes a single self-contained `dist-standalone/index.html` (CSS + views inlined). If the env var is omitted, add `<meta name="app-script-base" content="...">` or set `window.APP_RPC_BASE` manually before loading `util.html`.
3. Host `dist-standalone/index.html` on GitHub Pages / Netlify / etc. The client will:
   - prefer `google.script.run` when embedded inside Apps Script (unchanged behavior)
   - fall back to `fetch(<base>/exec)` elsewhere (requires the `base` from step 1)
4. For local dev, serve `dist-standalone/` (e.g., `npx http-server dist-standalone -p 5173`). Apps Script automatically returns `Access-Control-Allow-Origin: *`, so no extra CORS configuration is required. Use `text/plain` JSON payloads to avoid preflight checks.

Security tip: move toward a bearer token check in `rpc.ts` (stored in Script Properties) if the standalone site is public.
