import { getSongsWithLinksForView } from './features/songs';
import { listServices } from './features/services';
import { getFilesForFolderUrl } from './util/drive';
import { rpc } from './rpc';

export function doGet(e?: GoogleAppsScript.Events.DoGet) {
  const action = e?.parameter?.action;
  const viewMode = String(e?.parameter?.mode || e?.parameter?.view || '').toLowerCase();
  const guestMode = viewMode === 'guest';

  // JSON API: list files for a folder
  if (action === 'files') {
    const folderUrl = String(e?.parameter?.folderUrl || '');
    const files = getFilesForFolderUrl(folderUrl, 200);
    return ContentService
      .createTextOutput(JSON.stringify({ files }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // HTML app
  const tpl = HtmlService.createTemplateFromFile('index');
  tpl.rowsData = getSongsWithLinksForView();
  try { tpl.servicesData = listServices(); } catch (_) { tpl.servicesData = { items: [] }; }
  tpl.guestMode = guestMode;
  try {
    // Provide the deployed Web App base URL to client for fetch fallbacks
    // Note: returns null when not deployed as a Web App
    tpl.baseUrl = (ScriptApp.getService && ScriptApp.getService().getUrl && ScriptApp.getService().getUrl()) || '';
  } catch (_) {
    tpl.baseUrl = '';
  }
  return tpl.evaluate().setTitle('Worship Planner');
}

const withCors = (output: GoogleAppsScript.Content.TextOutput) => {
  const setter = (output as GoogleAppsScript.Content.TextOutput & { setHeader?: (key: string, value: string) => GoogleAppsScript.Content.TextOutput }).setHeader;
  if (typeof setter === 'function') {
    setter.call(output, 'Access-Control-Allow-Origin', '*');
    setter.call(output, 'Access-Control-Allow-Methods', 'POST,OPTIONS');
    setter.call(output, 'Access-Control-Allow-Headers', 'Content-Type,Authorization');
  }
  return output;
};

const emptyJson = (payload: unknown) =>
  withCors(
    ContentService
      .createTextOutput(JSON.stringify(payload ?? null))
      .setMimeType(ContentService.MimeType.JSON)
  );

export function doPost(e?: GoogleAppsScript.Events.DoPost) {
  try {
    const body = e?.postData?.contents || '';
    const parsed = body ? JSON.parse(body) : {};
    const method = String(parsed?.method || '');
    if (!method) throw new Error('Missing RPC method');
    const data = rpc({ method, payload: parsed?.payload });
    return emptyJson({ ok: true, data });
  } catch (err) {
    const message = err && (err as Error).message ? (err as Error).message : 'RPC failed';
    return emptyJson({ ok: false, error: message });
  }
}

export function doOptions() {
  return withCors(ContentService.createTextOutput(''));
}
