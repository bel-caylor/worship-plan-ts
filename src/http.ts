import { getSongsWithLinksForView } from './features/songs';
import { listServices } from './features/services';
import { getFilesForFolderUrl } from './util/drive';
import { getViewerProfile } from './features/roles';
import { rpc } from './rpc';
import { requestTokenEmail } from './auth';

export function doGet(e?: GoogleAppsScript.Events.DoGet) {
  const action = e?.parameter?.action;
  const viewMode = String(e?.parameter?.mode || e?.parameter?.view || '').toLowerCase();
  const viewerProfile = (() => {
    try { return getViewerProfile(); }
    catch (_) { return null; }
  })();
  const enforcedGuest = viewMode === 'guest';
  const defaultGuest = !(viewerProfile?.capabilities?.canEditPlan);
  const guestMode = enforcedGuest || defaultGuest;

  // JSON API: list files for a folder
  if (action === 'files') {
    const folderUrl = String(e?.parameter?.folderUrl || '');
    const files = getFilesForFolderUrl(folderUrl, 200);
    return ContentService
      .createTextOutput(JSON.stringify({ files }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (viewMode === 'login') {
    const tplLogin = HtmlService.createTemplateFromFile('login');
    return tplLogin.evaluate().setTitle('Admin Login');
  }

  // HTML app
  const tpl = HtmlService.createTemplateFromFile('index');
  tpl.rowsData = getSongsWithLinksForView();
  try { tpl.servicesData = listServices(); } catch (_) { tpl.servicesData = { items: [] }; }
  tpl.guestMode = guestMode;
  tpl.viewerProfile = viewerProfile;
  try {
    // Provide the deployed Web App base URL to client for fetch fallbacks
    // Note: returns null when not deployed as a Web App
    tpl.baseUrl = (ScriptApp.getService && ScriptApp.getService().getUrl && ScriptApp.getService().getUrl()) || '';
  } catch (_) {
    tpl.baseUrl = '';
  }
  return tpl.evaluate().setTitle('Worship Planner');
}

type MaybeHttpEvent = GoogleAppsScript.Events.DoGet | GoogleAppsScript.Events.DoPost | undefined;

const resolveOrigin = (event: MaybeHttpEvent): string => {
  const headerOrigin = (() => {
    const headers = (event as { headers?: { origin?: string } } | undefined)?.headers;
    return typeof headers?.origin === 'string' ? headers.origin : '';
  })();
  const paramOrigin = typeof event?.parameter?.origin === 'string' ? event?.parameter?.origin : '';
  return headerOrigin || paramOrigin || '';
};

const applyCors = (
  output: GoogleAppsScript.Content.TextOutput,
  origin?: string
) => {
  const setter = (output as GoogleAppsScript.Content.TextOutput & { setHeader?: (key: string, value: string) => GoogleAppsScript.Content.TextOutput }).setHeader;
  if (typeof setter === 'function') {
    const allowOrigin = origin && origin !== 'null' ? origin : '*';
    setter.call(output, 'Access-Control-Allow-Origin', allowOrigin);
    setter.call(output, 'Access-Control-Allow-Methods', 'POST,OPTIONS');
    setter.call(output, 'Access-Control-Allow-Headers', 'Content-Type,Authorization');
    setter.call(output, 'Vary', 'Origin');
    if (allowOrigin !== '*') {
      setter.call(output, 'Access-Control-Allow-Credentials', 'true');
    }
  }
  return output;
};

const makeJsonResponse = (payload: unknown, origin?: string) =>
  applyCors(
    ContentService
      .createTextOutput(JSON.stringify(payload ?? null))
      .setMimeType(ContentService.MimeType.JSON),
    origin
  );

export function doPost(e?: GoogleAppsScript.Events.DoPost) {
  const body = e?.postData?.contents || '';
  let parsed: { method?: string; payload?: unknown } = {};
  try { parsed = body ? JSON.parse(body) : {}; }
  catch (parseErr) {
    Logger.log(`doPost parse error: %s`, parseErr);
    return emptyJson({ ok: false, error: 'Invalid JSON payload' });
  }

  const origin = resolveOrigin(e);
  const headerAuth = (e as any)?.headers?.Authorization;
  const bearer = typeof headerAuth === 'string' ? headerAuth.replace(/^Bearer\s+/i, '').trim() : '';
  global.__REQUEST_AUTH_TOKEN__ = bearer || '';

  Logger.log(`doPost origin=%s method=%s`, origin || '???', parsed?.method || '');

  try {
    const method = String(parsed?.method || '');
    if (!method) throw new Error('Missing RPC method');
    const data = rpc({ method, payload: parsed?.payload });
    return makeJsonResponse({ ok: true, data }, origin);
  } catch (err) {
    const message = err && (err as Error).message ? (err as Error).message : 'RPC failed';
    return makeJsonResponse({ ok: false, error: message }, origin);
  } finally {
    global.__REQUEST_AUTH_TOKEN__ = '';
  }
}

export function doOptions(e?: GoogleAppsScript.Events.DoPost) {
  const origin = resolveOrigin(e);
  return makeJsonResponse('', origin);
}
