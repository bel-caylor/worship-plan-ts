Project conventions for agents and contributors

Overview
- This Apps Script project renders an HTML app with Alpine.js.
- Views contain only markup and bindings: `src/html/views/*.html`.
- Alpine app logic lives in head‑loaded script partials: `src/html/js/apps-*.html`.
- Client UI constants live in `src/html/context.html` (the "context pack").
- Server code (Apps Script) lives under `src/*.ts` and is bundled to `dist/Code.gs`.

Load order (index.html head)
1. styles
2. Alpine CDN
3. util.html (helpers shared by views)
4. context.html (window.APP_CTX with UI defaults/constants)
5. apps-songs.html and apps-weekly-plan.html (define songsApp/weeklyPlanApp)

Views
- Mount with `x-data="<appFn>()" x-init="init()"`.
- No inline `<script>` tags or closing `</body>` in view partials.

Client constants (context pack)
- `src/html/context.html` exports `window.APP_CTX`:
  - `defaults` (leader, sermon, serviceType, time)
  - `leaderChoices`, `sermonChoices`
  - `bibleBooks` as `[ [Book, chapterCount], ... ]`
- Alpine apps read from `APP_CTX` in `init()` and then merge any dynamic data (e.g., names discovered from Sheets).

RPC
- Client uses `google.script.run.rpc({ method, payload })` via `callRpc()` in `util.html`.
- Server routes inside `src/rpc.ts` with a `switch` on `method`.
- Add new RPCs by:
  1) Implementing a server function in `src/**.ts`.
  2) Adding a case in `src/rpc.ts` that calls it.
  3) Calling `callRpc('MethodName', payload)` from the client.

Build and deploy
- Build: `npm run build` (writes to `dist/`).
- Push:  `npx clasp push` (or `-f` to force).
- Apps Script deployment is performed by the user in the Apps Script UI: Deploy → Web app → New deployment.
- **Mandatory Cloudflare release handoff:** this app is served through the `worship-plan-proxy` Cloudflare Worker, not directly through an Apps Script URL. After the user creates an Apps Script deployment and provides its new `/exec` URL, update the Worker secret before calling the release complete:
  1. Set `APPS_SCRIPT_BASE` in `worship-plan-proxy` to that exact deployment base URL (the Worker appends `/exec`).
  2. Run `npm run build:standalone` from the repository root so the current client scripts are copied into `worship-plan-proxy/public`.
  3. Run `npm run deploy` from `worship-plan-proxy` to publish the Worker and its static assets.
  4. Verify a read-only RPC through `https://worship-plan-proxy.belinda-caylor.workers.dev` returns JSON with `ok: true`.
  5. Only then report the Cloudflare release as live and commit/push the release changes when requested.
- Never assume a previously deployed Apps Script URL remains the Worker upstream. Do not update an arbitrary historical Apps Script deployment: ask the user for the URL if it was not supplied.
- Manifest (`appsscript.json`) is in `dist/` via copy and controls runtime/scopes.

Editing rules for agents
- Prefer changing app logic in `src/html/js/apps-*.html`.
- Add UI constants in `src/html/context.html` (do not hardcode inside apps).
- Do not place secrets client‑side. Server secrets live only in Script Properties.
- When touching spreadsheet columns server‑side, use names from `src/constants.ts`.
