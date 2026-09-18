import { FOLDER_LINK_COL, MEDIA_FILES_COL, MEDIA_LAST_SYNCED_COL, SONG_SHEET } from '../constants';
import { getFilesForFolderUrl } from '../util/drive';
import { ensureColumn, getHeaders, getSheetByName } from '../util/sheets';

const REFRESH_CURSOR_KEY = 'songMediaRefreshCursor';
const MAX_ROWS_PER_RUN = 250;
const MAX_RUNTIME_MS = 4 * 60 * 1000;
type MediaFile = { name?: string; url?: string; mimeType?: string; path?: string };

/** Save Drive file links in Sheets so normal viewer requests never enumerate Drive. */
export function refreshSavedSongMedia() {
  const startedAt = Date.now();
  const sheet = getSheetByName(SONG_SHEET);
  const { headers, colMap } = getHeaders(sheet);
  ensureColumn(sheet, headers, colMap, FOLDER_LINK_COL);
  ensureColumn(sheet, headers, colMap, MEDIA_FILES_COL);
  ensureColumn(sheet, headers, colMap, MEDIA_LAST_SYNCED_COL);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { updated: 0, remaining: false };
  const folderIndex = colMap[FOLDER_LINK_COL];
  const filesColumn = colMap[MEDIA_FILES_COL] + 1;
  const syncedColumn = colMap[MEDIA_LAST_SYNCED_COL] + 1;
  const properties = PropertiesService.getScriptProperties();
  const storedCursor = Number(properties.getProperty(REFRESH_CURSOR_KEY)) || 2;
  let rowNumber = Math.min(Math.max(storedCursor, 2), lastRow);
  let updated = 0;
  let inspected = 0;

  while (rowNumber <= lastRow && inspected < MAX_ROWS_PER_RUN && Date.now() - startedAt < MAX_RUNTIME_MS) {
    const range = sheet.getRange(rowNumber, 1, 1, headers.length);
    const values = range.getValues()[0];
    const formulas = range.getFormulas()[0];
    const rich = range.getRichTextValues()[0];
    const folderUrl = folderUrlFromCell(values[folderIndex], formulas[folderIndex], rich[folderIndex]);
    inspected += 1;
    if (folderUrl) {
      try {
        writeMediaSnapshot(sheet.getRange(rowNumber, filesColumn), getFilesForFolderUrl(folderUrl, 60));
        sheet.getRange(rowNumber, syncedColumn).setValue(new Date());
        updated += 1;
      } catch (err) {
        // Keep the prior snapshot if Drive is temporarily unavailable.
        Logger.log(`Saved media refresh failed for Songs row ${rowNumber}: ${err}`);
      }
    }
    rowNumber += 1;
  }
  const remaining = rowNumber <= lastRow;
  if (remaining) properties.setProperty(REFRESH_CURSOR_KEY, String(rowNumber));
  else properties.deleteProperty(REFRESH_CURSOR_KEY);
  try {
    SpreadsheetApp.getActive().toast(
      remaining ? `Saved links refreshed for ${updated} songs; the next run will continue.` : `Saved links refreshed for ${updated} songs.`,
      'Worship Planner',
      5
    );
  } catch (_) { /* Time triggers have no visible spreadsheet UI. */ }
  return { updated, remaining };
}

/** Install exactly one overnight schedule; repeated calls are safe. */
export function installDailySongMediaRefresh() {
  const handler = 'refreshSavedSongMedia';
  const installed = ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === handler);
  if (!installed) ScriptApp.newTrigger(handler).timeBased().everyDays(1).atHour(3).create();
  SpreadsheetApp.getActive().toast(installed ? 'Daily media refresh is already installed.' : 'Daily media refresh installed.', 'Worship Planner', 4);
  return { installed: !installed };
}

export function readMediaSnapshot(richText: GoogleAppsScript.Spreadsheet.RichTextValue | null): MediaFile[] {
  if (!richText) return [];
  const files: MediaFile[] = [];
  try {
    for (const run of richText.getRuns() || []) {
      const url = String(run.getLinkUrl() || '').trim();
      const name = String(run.getText() || '').trim();
      if (url && name) files.push({ name, url, mimeType: '', path: '' });
    }
    if (!files.length) {
      const url = String(richText.getLinkUrl() || '').trim();
      const name = String(richText.getText() || '').trim();
      if (url && name) files.push({ name, url, mimeType: '', path: '' });
    }
  } catch (_) { /* Treat a malformed rich-text cell as an empty snapshot. */ }
  return files;
}

function writeMediaSnapshot(cell: GoogleAppsScript.Spreadsheet.Range, files: MediaFile[]) {
  const usable = (Array.isArray(files) ? files : []).filter(file => String(file?.name || '').trim() && String(file?.url || '').trim()).slice(0, 60);
  if (!usable.length) { cell.clearContent(); return; }
  let text = '';
  const links: Array<{ start: number; end: number; url: string }> = [];
  usable.forEach((file, index) => {
    const name = String(file.name).trim();
    const start = text.length;
    text += name + (index < usable.length - 1 ? '\n' : '');
    links.push({ start, end: start + name.length, url: String(file.url).trim() });
  });
  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  links.forEach(link => builder.setLinkUrl(link.start, link.end, link.url));
  cell.setRichTextValue(builder.build());
  cell.setWrap(true);
}

function folderUrlFromCell(value: unknown, formula: string, richText: GoogleAppsScript.Spreadsheet.RichTextValue | null) {
  try {
    const direct = richText?.getLinkUrl();
    if (direct) return direct;
    for (const run of richText?.getRuns() || []) { const link = run.getLinkUrl(); if (link) return link; }
  } catch (_) { /* Fall through to formula and displayed value. */ }
  const formulaMatch = /^=HYPERLINK\("([^"]+)"/i.exec(String(formula || ''));
  if (formulaMatch) return formulaMatch[1];
  return String(value ?? '').match(/https:\/\/drive\.google\.com\/[^\s"]+/)?.[0] || '';
}
