type Env = {
  APPS_SCRIPT_BASE?: string;
  ASSETS?: Fetcher;
};

function normalizeAppsScriptBase(value?: string) {
  return String(value || '')
    .trim()
    .replace(/\/+$/, '')
    .replace(/\/(exec|dev)$/i, '');
}

export default {
  async fetch(request: Request, env: Env) {
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
    const appsScriptBase = normalizeAppsScriptBase(env.APPS_SCRIPT_BASE);
    if (!appsScriptBase) {
      return jsonError(
        origin,
        500,
        'Worker secret APPS_SCRIPT_BASE is not set. Point it at the current Apps Script web app deployment URL and redeploy the worker.'
      );
    }

    const rpcMethod = (() => {
      try { return String(JSON.parse(body || '{}')?.method || ''); }
      catch (_) { return ''; }
    })();
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
            upstream = await fetch(resultUrl, { method: 'GET', headers: { 'Cache-Control': 'no-store' } });
            const isJsonResult = upstream.ok && parseJsonSafely(await upstream.clone().text()).ok;
            if (isJsonResult) break;
            if (resultAttempt < 2) {
              await new Promise(resolve => setTimeout(resolve, 250 * (resultAttempt + 1)));
            }
          }
        }

        text = await upstream.text();
        contentType = upstream.headers.get('Content-Type') || '';
        parsed = parseJsonSafely(text);
        if (parsed.ok) break;
        // A brief backoff gives Google's one-time result URL time to become
        // available instead of returning its transient Drive 404 to the app.
        if (attempt < attempts - 1) {
          await new Promise(resolve => setTimeout(resolve, 250 * (attempt + 1)));
        }
      }
    } catch (err) {
      return jsonError(
        origin,
        502,
        `Unable to reach Apps Script: ${err instanceof Error ? err.message : String(err)}`
      );
    }

    if (!parsed.ok) {
      const preview = summarizeUpstream(text);
      return jsonError(
        origin,
        502,
        `Apps Script returned a non-JSON response (${upstream?.status || 502}${contentType ? `, ${contentType}` : ''}). The Worker may be pointed at an outdated/non-public deployment, or Apps Script may have returned a transient HTML response.`,
        preview
      );
    }

    return new Response(JSON.stringify(parsed.value), {
      status: upstream?.status || 200,
      headers: {
        ...cors(origin),
        'Content-Type': 'application/json; charset=utf-8'
      }
    });
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

function jsonError(origin: string, status: number, message: string, details?: string) {
  return new Response(JSON.stringify({
    ok: false,
    error: details ? `${message} Upstream preview: ${details}` : message
  }), {
    status,
    headers: {
      ...cors(origin),
      'Content-Type': 'application/json; charset=utf-8'
    }
  });
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
