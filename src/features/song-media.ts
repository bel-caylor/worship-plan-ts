import { FOLDER_LINK_COL, SONG_COL_NAME, SONG_SHEET } from '../constants';
import { getSheetByName } from '../util/sheets';

const folderIdFromUrl = (url: string) =>
  /\/folders\/([A-Za-z0-9_-]+)/.test(url) || /[?&]id=([A-Za-z0-9_-]+)/.test(url);

const urlFromCell = (value: unknown, formula: string, richText: GoogleAppsScript.Spreadsheet.RichTextValue | null) => {
  try {
    const direct = richText?.getLinkUrl();
    if (direct && folderIdFromUrl(direct)) return direct;
    for (const run of richText?.getRuns() || []) {
      const link = run.getLinkUrl();
      if (link && folderIdFromUrl(link)) return link;
    }
  } catch (_) { /* fall through to the formula and displayed value */ }

  for (const candidate of [formula, String(value ?? '')]) {
    const url = String(candidate || '').match(/https?:\/\/[^\s"')]+/i)?.[0] || '';
    if (url && folderIdFromUrl(url)) return url;
  }
  return '';
};

/** Resolves a song's Drive folder when an older catalog payload lacks _folderUrl. */
export function getSongFolderUrl(input: { songName?: string }) {
  const songName = String(input?.songName || '').trim().toLowerCase();
  if (!songName) return '';

  const sheet = getSheetByName(SONG_SHEET);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || !lastColumn) return '';

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(String);
  const nameColumn = headers.findIndex(header => header.trim() === SONG_COL_NAME);
  const folderColumn = headers.findIndex(header => header.trim() === FOLDER_LINK_COL);
  if (nameColumn < 0 || folderColumn < 0) return '';

  const range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const richText = range.getRichTextValues();
  const row = values.findIndex(item => String(item[nameColumn] ?? '').trim().toLowerCase() === songName);
  if (row < 0) return '';

  return urlFromCell(values[row][folderColumn], formulas[row][folderColumn], richText[row][folderColumn]);
}
