import { getSongsWithLinksForView } from './features/songs';
import { getFilesForFolderUrl } from './util/drive';

export function doGet(e?: GoogleAppsScript.Events.DoGet) {
  const action = e?.parameter?.action;

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
  try {
    // Provide the deployed Web App base URL to client for fetch fallbacks
    // Note: returns null when not deployed as a Web App
    tpl.baseUrl = (ScriptApp.getService && ScriptApp.getService().getUrl && ScriptApp.getService().getUrl()) || '';
  } catch (_) {
    tpl.baseUrl = '';
  }
  return tpl.evaluate().setTitle('Worship Planner');
}
