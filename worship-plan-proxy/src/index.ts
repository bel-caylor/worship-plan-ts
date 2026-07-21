type Env = {
  APPS_SCRIPT_BASE?: string;
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

    const upstream = await fetch(`${appsScriptBase}/exec`, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body
    });

    const text = await upstream.text();
    const contentType = upstream.headers.get('Content-Type') || '';
    const parsed = parseJsonSafely(text);
    if (!parsed.ok) {
      const preview = summarizeUpstream(text);
      return jsonError(
        origin,
        502,
        `Apps Script returned a non-JSON response (${upstream.status}${contentType ? `, ${contentType}` : ''}). This usually means the worker is pointing at an outdated or non-public Apps Script deployment.`,
        preview
      );
    }

    return new Response(JSON.stringify(parsed.value), {
      status: upstream.status,
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

function cors(origin) {
  const allow = origin && origin !== 'null' ? origin : '*';
  return {
    'Access-Control-Allow-Origin': allow,
    'Access-Control-Allow-Methods': 'POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type,Authorization',
    'Vary': 'Origin'
  };
}
