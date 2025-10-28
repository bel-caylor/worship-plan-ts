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
  return tpl.evaluate().setTitle('Worship Planner');
}
