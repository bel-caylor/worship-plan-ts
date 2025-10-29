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
- Deploy via Apps Script UI: Deploy → Web app → New deployment.
- Manifest (`appsscript.json`) is in `dist/` via copy and controls runtime/scopes.

Editing rules for agents
- Prefer changing app logic in `src/html/js/apps-*.html`.
- Add UI constants in `src/html/context.html` (do not hardcode inside apps).
- Do not place secrets client‑side. Server secrets live only in Script Properties.
- When touching spreadsheet columns server‑side, use names from `src/constants.ts`.

