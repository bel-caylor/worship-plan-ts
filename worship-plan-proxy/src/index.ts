type Env = {
  APPS_SCRIPT_BASE?: string;
};

const DEFAULT_APPS_SCRIPT_BASE = 'https://script.google.com/macros/s/AKfycbxWN6rc4JgBBiUzS3QXTAESUK_lDsv7VgX9H6RObAxABS1o8qVQ6MkciUSdBxGnNSoM';

function normalizeAppsScriptBase(value?: string) {
  return String(value || DEFAULT_APPS_SCRIPT_BASE)
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

    const body = await request.text();
    const appsScriptBase = normalizeAppsScriptBase(env.APPS_SCRIPT_BASE);
    const upstream = await fetch(`${appsScriptBase}/exec`, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body
    });

    const text = await upstream.text();
    return new Response(text, {
      status: upstream.status,
      headers: {
        ...cors(origin),
        'Content-Type': upstream.headers.get('Content-Type') || 'application/json'
      }
    });
  }
};

function cors(origin) {
  const allow = origin && origin !== 'null' ? origin : '*';
  return {
    'Access-Control-Allow-Origin': allow,
    'Access-Control-Allow-Methods': 'POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type,Authorization',
    'Vary': 'Origin'
  };
}
