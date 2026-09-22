type Env = {
  APPS_SCRIPT_BASE?: string;
  ASSETS?: Fetcher;
};

const EDGE_CACHEABLE_RPC_TTL_SECONDS: Record<string, number> = {
  getServiceViewerStartup: 120
};

function normalizeAppsScriptBase(value?: string) {
  return String(value || '')
    .trim()
    .replace(/\/+$/, '')
    .replace(/\/(exec|dev)$/i, '');
}

export default {
  async fetch(request: Request, env: Env) {
    const workerStartedAt = Date.now();
    const phases: Array<Record<string, unknown>> = [];
    const markPhase = (name: string, startedAt: number, detail: Record<string, unknown> = {}) => {
      const elapsedMs = Date.now() - startedAt;
      phases.push({
        name,
        elapsedMs,
        seconds: Number((elapsedMs / 1000).toFixed(2)),
        ...detail
      });
    };
    const origin = request.headers.get('Origin') || '';
    if (request.method === 'OPTIONS') {
      return new Response('', { headers: cors(origin) });
    }

    if (request.method !== 'POST') {
      if (env.ASSETS) return env.ASSETS.fetch(request);
      return new Response(
        'Worship Plan Proxy is running. Send POST RPC requests from the app to this URL.',
        {
          status: 200,
          headers: {
            ...cors(origin),
            'Content-Type': 'text/plain; charset=utf-8'
          }
        }
      );
    }

    const body = await request.text();
    let parsedBody: { method?: string; payload?: unknown; debugRpc?: unknown } = {};
    try {
      parsedBody = JSON.parse(body || '{}');
    } catch (_) {
      parsedBody = {};
    }
    const debugRpc = parsedBody?.debugRpc === true;
    const rpcMethod = String(parsedBody?.method || '');
    const appsScriptBase = normalizeAppsScriptBase(env.APPS_SCRIPT_BASE);
    if (!appsScriptBase) {
      return jsonError(
        origin,
        500,
        'Worker secret APPS_SCRIPT_BASE is not set. Point it at the current Apps Script web app deployment URL and redeploy the worker.',
        undefined,
        debugRpc ? workerDebug(rpcMethod, workerStartedAt, phases, 'config_error') : undefined
      );
    }

    const edgeCacheTtl = EDGE_CACHEABLE_RPC_TTL_SECONDS[rpcMethod] || 0;
    const edgeCacheKey = edgeCacheTtl ? await rpcEdgeCacheKey(request, body) : null;
    if (edgeCacheKey) {
      const cacheStartedAt = Date.now();
      const cached = await caches.default.match(edgeCacheKey);
      markPhase('edge_cache_match', cacheStartedAt, { hit: Boolean(cached) });
      if (cached) {
        return withCors(cached, origin);
      }
    }
    // Google intermittently returns a Drive 404 from its Apps Script redirect
    // path when requests originate at the Worker. Retrying reads is safe; so
    // is saveOrder because it replaces the complete order with the same body.
    // Never retry email, create, or other potentially non-idempotent writes.
    const canRetry = /^(get|list|suggest|ai|summarize|esv)/i.test(rpcMethod)
      || rpcMethod === 'memberExistsInRoles'
      || rpcMethod === 'saveOrder';

    let upstream: Response | undefined;
    let text = '';
    let contentType = '';
    let parsed: ReturnType<typeof parseJsonSafely> = { ok: false };
    try {
      // A retry is enough to cover the occasional Google redirect race. Four
      // complete POST cycles can turn one bad read into a multi-second hang.
      const attempts = canRetry ? 2 : 1;
      for (let attempt = 0; attempt < attempts; attempt += 1) {
        const rpcUrl = `${appsScriptBase}/exec?worker_request=${Date.now()}_${attempt}`;
        const postStartedAt = Date.now();
        upstream = await fetch(rpcUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'text/plain;charset=utf-8',
            'Cache-Control': 'no-store'
          },
          body,
          // Apps Script answers a POST with a 302 to a one-time
          // googleusercontent URL. Retrieve that response explicitly as GET.
          redirect: 'manual'
        });
        markPhase('apps_script_post', postStartedAt, {
          attempt,
          status: upstream.status
        });

        if (upstream.status >= 300 && upstream.status < 400) {
          const resultUrl = upstream.headers.get('Location');
          if (!resultUrl) throw new Error('Apps Script redirected the RPC request without a result URL.');
          // The redirected googleusercontent URL is occasionally not ready
          // immediately, especially just after a new Apps Script deployment.
          // Retrying this same one-time URL is important; issuing a new POST
          // only creates another result URL that has the same race. A 200 can
          // still be Google's temporary HTML page, so check the body before
          // treating the result as ready. This is safe for email sends because
          // it never replays their original POST.
          for (let resultAttempt = 0; resultAttempt < 3; resultAttempt += 1) {
            const resultStartedAt = Date.now();
            upstream = await fetch(resultUrl, { method: 'GET', headers: { 'Cache-Control': 'no-store' } });
            const probeText = await upstream.clone().text();
            const isJsonResult = upstream.ok && parseJsonSafely(probeText).ok;
            markPhase('apps_script_result_get', resultStartedAt, {
              attempt,
              resultAttempt,
              status: upstream.status,
              isJson: isJsonResult,
              bytes: probeText.length
            });
            if (isJsonResult) break;
            if (resultAttempt < 2) {
              const backoffStartedAt = Date.now();
              await new Promise(resolve => setTimeout(resolve, 250 * (resultAttempt + 1)));
              markPhase('result_backoff', backoffStartedAt, { attempt, resultAttempt });
            }
          }
        }

        const readStartedAt = Date.now();
        text = await upstream.text();
        contentType = upstream.headers.get('Content-Type') || '';
        parsed = parseJsonSafely(text);
        markPhase('upstream_read_parse', readStartedAt, {
          attempt,
          status: upstream.status,
          contentType,
          isJson: parsed.ok,
          bytes: text.length
        });
        if (parsed.ok) break;
        // A brief backoff gives Google's one-time result URL time to become
        // available instead of returning its transient Drive 404 to the app.
        if (attempt < attempts - 1) {
          const backoffStartedAt = Date.now();
          await new Promise(resolve => setTimeout(resolve, 250 * (attempt + 1)));
          markPhase('retry_backoff', backoffStartedAt, { attempt });
        }
      }
    } catch (err) {
      return jsonError(
        origin,
        502,
        `Unable to reach Apps Script: ${err instanceof Error ? err.message : String(err)}`,
        undefined,
        debugRpc ? workerDebug(rpcMethod, workerStartedAt, phases, 'exception') : undefined
      );
    }

    if (!parsed.ok) {
      const preview = summarizeUpstream(text);
      return jsonError(
        origin,
        502,
        `Apps Script returned a non-JSON response (${upstream?.status || 502}${contentType ? `, ${contentType}` : ''}). The Worker may be pointed at an outdated/non-public deployment, or Apps Script may have returned a transient HTML response.`,
        preview,
        debugRpc ? workerDebug(rpcMethod, workerStartedAt, phases, 'non_json') : undefined
      );
    }

    const responsePayload = debugRpc && parsed.value && typeof parsed.value === 'object'
      ? { ...(parsed.value as Record<string, unknown>), __debug: workerDebug(rpcMethod, workerStartedAt, phases, 'ok') }
      : parsed.value;
    const response = new Response(JSON.stringify(responsePayload), {
      status: upstream?.status || 200,
      headers: {
        ...cors(origin),
        'Content-Type': 'application/json; charset=utf-8'
      }
    });
    if (edgeCacheKey && edgeCacheTtl > 0 && !debugRpc) {
      const cacheResponse = new Response(response.clone().body, response);
      cacheResponse.headers.set('Cache-Control', `public, max-age=${edgeCacheTtl}`);
      const putStartedAt = Date.now();
      await caches.default.put(edgeCacheKey, cacheResponse);
      markPhase('edge_cache_put', putStartedAt);
    }
    return response;
  }
};

function parseJsonSafely(text: string) {
  try {
    return { ok: true as const, value: JSON.parse(text) };
  } catch (_) {
    return { ok: false as const };
  }
}

function summarizeUpstream(text: string) {
  return String(text || '')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 220);
}

function jsonError(origin: string, status: number, message: string, details?: string, debug?: Record<string, unknown>) {
  return new Response(JSON.stringify({
    ok: false,
    error: details ? `${message} Upstream preview: ${details}` : message,
    ...(debug ? { __debug: debug } : {})
  }), {
    status,
    headers: {
      ...cors(origin),
      'Content-Type': 'application/json; charset=utf-8'
    }
  });
}

function workerDebug(method: string, startedAt: number, phases: Array<Record<string, unknown>>, outcome: string) {
  const elapsedMs = Date.now() - startedAt;
  return {
    method,
    outcome,
    elapsedMs,
    seconds: Number((elapsedMs / 1000).toFixed(2)),
    phases
  };
}

async function rpcEdgeCacheKey(request: Request, body: string) {
  let parsed: { method?: string; payload?: unknown } = {};
  try {
    parsed = JSON.parse(body || '{}');
  } catch (_) {
    parsed = {};
  }
  const method = String(parsed?.method || '');
  if (!EDGE_CACHEABLE_RPC_TTL_SECONDS[method]) return null;
  const payload = JSON.stringify(parsed?.payload ?? null);
  const digest = await sha256Hex(`${method}|${payload}`);
  return new Request(new URL(`/__rpc_cache/${method}/${digest}`, request.url).toString(), {
    method: 'GET'
  });
}

async function sha256Hex(value: string) {
  const bytes = new TextEncoder().encode(value);
  const hash = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(hash))
    .map(byte => byte.toString(16).padStart(2, '0'))
    .join('');
}

function withCors(response: Response, origin: string) {
  const next = new Response(response.body, response);
  Object.entries(cors(origin)).forEach(([key, value]) => {
    next.headers.set(key, value);
  });
  next.headers.set('Content-Type', 'application/json; charset=utf-8');
  return next;
}

function cors(origin: string) {
  const allow = origin && origin !== 'null' ? origin : '*';
  return {
    'Access-Control-Allow-Origin': allow,
    'Access-Control-Allow-Methods': 'POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type,Authorization',
    'Vary': 'Origin'
  };
}
