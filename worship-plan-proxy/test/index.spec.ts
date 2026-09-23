import { createExecutionContext, waitOnExecutionContext } from 'cloudflare:test';
import { afterEach, describe, expect, it, vi } from 'vitest';
import worker from '../src';

const env = {
	APPS_SCRIPT_BASE: 'https://script.google.com/macros/s/test-deployment/exec'
};

function rpcRequest(method = 'getServiceViewerStartup', payload: unknown = null) {
	return new Request('https://proxy.example.test/exec', {
		method: 'POST',
		headers: {
			'Content-Type': 'text/plain;charset=utf-8',
			Origin: 'https://app.example.test'
		},
		body: JSON.stringify({ method, payload })
	});
}

async function fetchWorker(request: Request, testEnv = env) {
	const ctx = createExecutionContext();
	const response = await worker.fetch(request, testEnv, ctx);
	await waitOnExecutionContext(ctx);
	return response;
}

afterEach(() => {
	vi.unstubAllGlobals();
});

describe('Worship Plan proxy RPC contract', () => {
	it('always returns JSON when APPS_SCRIPT_BASE is missing', async () => {
		const response = await fetchWorker(rpcRequest(), {});
		const body = await response.json() as { ok: boolean; error: string };

		expect(response.status).toBe(500);
		expect(response.headers.get('Content-Type')).toContain('application/json');
		expect(body.ok).toBe(false);
		expect(body.error).toContain('APPS_SCRIPT_BASE is not set');
	});

	it('forwards direct JSON Apps Script responses as JSON', async () => {
		const fetchMock = vi.fn().mockResolvedValue(
			new Response(JSON.stringify({ ok: true, data: { ready: true } }), {
				status: 200,
				headers: { 'Content-Type': 'application/json; charset=utf-8' }
			})
		);
		vi.stubGlobal('fetch', fetchMock);

		const response = await fetchWorker(rpcRequest('listServices', { summary: true }));
		const body = await response.json() as { ok: boolean; data: { ready: boolean } };

		expect(response.status).toBe(200);
		expect(response.headers.get('Content-Type')).toContain('application/json');
		expect(body).toEqual({ ok: true, data: { ready: true } });
		expect(fetchMock).toHaveBeenCalledTimes(1);
		expect(fetchMock.mock.calls[0][0]).toContain('/exec?worker_request=');
		expect(fetchMock.mock.calls[0][1]).toMatchObject({
			method: 'POST',
			redirect: 'manual'
		});
	});

	it('follows Apps Script redirects and retries temporary non-JSON result pages', async () => {
		const fetchMock = vi.fn()
			.mockResolvedValueOnce(new Response('', {
				status: 302,
				headers: { Location: 'https://script.googleusercontent.com/result' }
			}))
			.mockResolvedValueOnce(new Response('<html>Temporary Drive 404</html>', {
				status: 200,
				headers: { 'Content-Type': 'text/html; charset=utf-8' }
			}))
			.mockResolvedValueOnce(new Response(JSON.stringify({ ok: true, data: { service: null, items: [] } }), {
				status: 200,
				headers: { 'Content-Type': 'application/json; charset=utf-8' }
			}));
		vi.stubGlobal('fetch', fetchMock);

		const response = await fetchWorker(rpcRequest('getServiceViewerStartup'));
		const body = await response.json() as { ok: boolean; data: { items: unknown[] } };

		expect(response.status).toBe(200);
		expect(body.ok).toBe(true);
		expect(body.data.items).toEqual([]);
		expect(fetchMock).toHaveBeenCalledTimes(3);
		expect(fetchMock.mock.calls[1][0]).toBe('https://script.googleusercontent.com/result');
		expect(fetchMock.mock.calls[2][0]).toBe('https://script.googleusercontent.com/result');
	});

	it('wraps persistent non-JSON upstream responses in a JSON error', async () => {
		const fetchMock = vi.fn()
			.mockResolvedValueOnce(new Response('', {
				status: 302,
				headers: { Location: 'https://script.googleusercontent.com/result' }
			}))
			.mockImplementation(() => Promise.resolve(
				new Response('<html><title>Drive 404</title><p>Not ready</p></html>', {
					status: 404,
					headers: { 'Content-Type': 'text/html; charset=utf-8' }
				})
			));
		vi.stubGlobal('fetch', fetchMock);

		const response = await fetchWorker(rpcRequest('getOrder', '2026-09-22'));
		const body = await response.json() as { ok: boolean; error: string };

		expect(response.status).toBe(502);
		expect(response.headers.get('Content-Type')).toContain('application/json');
		expect(body.ok).toBe(false);
		expect(body.error).toContain('Apps Script returned a non-JSON response');
		expect(body.error).toContain('Drive 404');
	});

	it('retries the RPC immediately when a redirected result returns the app shell', async () => {
		const appShell = `<!doctype html><html><head><title>Worship Planner</title></head><body>${'x'.repeat(100001)}</body></html>`;
		const fetchMock = vi.fn()
			.mockResolvedValueOnce(new Response('', {
				status: 302,
				headers: { Location: 'https://script.googleusercontent.com/app-shell' }
			}))
			.mockResolvedValueOnce(new Response(appShell, {
				status: 200,
				headers: { 'Content-Type': 'text/html; charset=utf-8' }
			}))
			.mockResolvedValueOnce(new Response('', {
				status: 302,
				headers: { Location: 'https://script.googleusercontent.com/rpc-result' }
			}))
			.mockResolvedValueOnce(new Response(JSON.stringify({ ok: true, data: { items: [] } }), {
				status: 200,
				headers: { 'Content-Type': 'application/json; charset=utf-8' }
			}));
		vi.stubGlobal('fetch', fetchMock);

		const response = await fetchWorker(rpcRequest('listServices', { summary: true }));
		const body = await response.json() as { ok: boolean; data: { items: unknown[] } };

		expect(response.status).toBe(200);
		expect(body.ok).toBe(true);
		expect(body.data.items).toEqual([]);
		expect(fetchMock).toHaveBeenCalledTimes(4);
		expect(fetchMock.mock.calls[0][1]).toMatchObject({ method: 'POST' });
		expect(fetchMock.mock.calls[1][0]).toBe('https://script.googleusercontent.com/app-shell');
		expect(fetchMock.mock.calls[2][1]).toMatchObject({ method: 'POST' });
		expect(fetchMock.mock.calls[3][0]).toBe('https://script.googleusercontent.com/rpc-result');
	});

	it('retries safe RPCs after a redirected result times out', async () => {
		const fetchMock = vi.fn()
			.mockResolvedValueOnce(new Response('', {
				status: 302,
				headers: { Location: 'https://script.googleusercontent.com/slow-result' }
			}))
			.mockRejectedValueOnce(new DOMException('The operation was aborted.', 'AbortError'))
			.mockResolvedValueOnce(new Response('', {
				status: 302,
				headers: { Location: 'https://script.googleusercontent.com/rpc-result' }
			}))
			.mockResolvedValueOnce(new Response(JSON.stringify({ ok: true, data: { items: [] } }), {
				status: 200,
				headers: { 'Content-Type': 'application/json; charset=utf-8' }
			}));
		vi.stubGlobal('fetch', fetchMock);

		const response = await fetchWorker(rpcRequest('getViewerProfile'));
		const body = await response.json() as { ok: boolean; data: { items: unknown[] } };

		expect(response.status).toBe(200);
		expect(body.ok).toBe(true);
		expect(fetchMock).toHaveBeenCalledTimes(4);
		expect(fetchMock.mock.calls.filter(([, init]) => init?.method === 'POST')).toHaveLength(2);
	});

	it('does not report the original Apps Script redirect when the final redirected result times out', async () => {
		const fetchMock = vi.fn()
			.mockResolvedValueOnce(new Response('', {
				status: 302,
				headers: {
					Location: 'https://script.googleusercontent.com/slow-result-1',
					'Content-Type': 'application/binary'
				}
			}))
			.mockRejectedValueOnce(new DOMException('The operation was aborted.', 'AbortError'))
			.mockResolvedValueOnce(new Response('', {
				status: 302,
				headers: {
					Location: 'https://script.googleusercontent.com/slow-result-2',
					'Content-Type': 'application/binary'
				}
			}))
			.mockRejectedValueOnce(new DOMException('The operation was aborted.', 'AbortError'));
		vi.stubGlobal('fetch', fetchMock);

		const response = await fetchWorker(rpcRequest('getServiceTeamAssignments', { serviceId: 'svc-1' }));
		const body = await response.json() as { ok: boolean; error: string };

		expect(response.status).toBe(502);
		expect(body.ok).toBe(false);
		expect(body.error).toContain('upstream_timeout');
		expect(body.error).not.toContain('non-JSON response (302');
		expect(fetchMock.mock.calls.filter(([, init]) => init?.method === 'POST')).toHaveLength(2);
	});

	it('does not replay non-idempotent email RPC POSTs when the redirected result is non-JSON', async () => {
		const fetchMock = vi.fn()
			.mockResolvedValueOnce(new Response('', {
				status: 302,
				headers: { Location: 'https://script.googleusercontent.com/email-result' }
			}))
			.mockResolvedValue(new Response('<html>temporary email result page</html>', {
				status: 200,
				headers: { 'Content-Type': 'text/html; charset=utf-8' }
			}));
		vi.stubGlobal('fetch', fetchMock);

		const response = await fetchWorker(rpcRequest('sendServiceTeamEmail', { serviceId: 'svc-1' }));
		const body = await response.json() as { ok: boolean };

		expect(response.status).toBe(502);
		expect(body.ok).toBe(false);
		expect(fetchMock.mock.calls.filter(([, init]) => init?.method === 'POST')).toHaveLength(1);
		expect(fetchMock.mock.calls.filter(([, init]) => init?.method === 'GET')).toHaveLength(3);
	});
});
