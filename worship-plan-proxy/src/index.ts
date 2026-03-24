const APPS_SCRIPT_BASE = 'https://script.google.com/macros/s/AKfycbzLVHRDLYH49I9CIsmem1jaZHgFjQJaBX_X4QKfpnaYskKSvYPeUYK6uf5xR24gzlTa'; // no /exec

export default {
  async fetch(request) {
    const origin = request.headers.get('Origin') || '';
    if (request.method === 'OPTIONS') {
      return new Response('', { headers: cors(origin) });
    }

    const body = await request.text();
    const upstream = await fetch(`${APPS_SCRIPT_BASE}/exec`, {
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
