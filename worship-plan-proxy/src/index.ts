type Env = {
  APPS_SCRIPT_BASE?: string;
};

const DEFAULT_APPS_SCRIPT_BASE = 'https://script.google.com/macros/s/AKfycbyJjOwVSDYeaiJxAjWSKZePBD8BK9_fmoKDvCB_XcJaTUxnGU6D0YUbf9fsdWnZZjgv';

export default {
  async fetch(request: Request, env: Env) {
    const origin = request.headers.get('Origin') || '';
    if (request.method === 'OPTIONS') {
      return new Response('', { headers: cors(origin) });
    }

    const body = await request.text();
    const appsScriptBase = String(env.APPS_SCRIPT_BASE || DEFAULT_APPS_SCRIPT_BASE).replace(/\/+$/, '');
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
