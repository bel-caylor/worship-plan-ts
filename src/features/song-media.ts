import { FOLDER_LINK_COL, SONG_COL_NAME, SONG_SHEET } from '../constants';
import { getSheetByName } from '../util/sheets';
import { normalize, similarity } from '../util/text';

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

/**
 * Resolves a song's Drive folder when the catalog payload or its sheet row has
 * no link yet.  The Drive fallback keeps media available before the admin runs
 * the bulk "Link song media" command.
 */
export function getSongFolderUrl(input: { songName?: string }) {
  const rawSongName = String(input?.songName || '').trim();
  const songName = rawSongName.toLowerCase();
  if (!rawSongName) return '';

  const sheet = getSheetByName(SONG_SHEET);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || !lastColumn) return findSongFolderInDrive(rawSongName);

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(String);
  const nameColumn = headers.findIndex(header => header.trim() === SONG_COL_NAME);
  const folderColumn = headers.findIndex(header => header.trim() === FOLDER_LINK_COL);
  if (nameColumn < 0 || folderColumn < 0) return findSongFolderInDrive(rawSongName);

  const range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const richText = range.getRichTextValues();
  const row = values.findIndex(item => String(item[nameColumn] ?? '').trim().toLowerCase() === songName);
  if (row < 0) return findSongFolderInDrive(rawSongName);

  return urlFromCell(values[row][folderColumn], formulas[row][folderColumn], richText[row][folderColumn])
    || findSongFolderInDrive(rawSongName);
}

function findSongFolderInDrive(songName: string) {
  const expectedName = normalize(songName);
  if (!expectedName) return '';
  try {
    // Drive performs title searches server-side. Walking every folder under
    // the music library here can exceed a web-app request timeout.
    const terms = [songName, ...expectedName.split(' ').filter(token => token.length >= 4)]
      .filter((term, index, all) => all.indexOf(term) === index);
    let best: { url: string; score: number } | null = null;
    for (const term of terms.slice(0, 4)) {
      const escapedTerm = term.replace(/\\/g, '\\\\').replace(/'/g, "\\\\'");
      const folders = DriveApp.searchFolders(`title contains '${escapedTerm}' and trashed = false`);
      while (folders.hasNext()) {
        const folder = folders.next();
        const score = similarity(normalize(folder.getName()), expectedName);
        if (!best || score > best.score) best = { url: folder.getUrl(), score };
      }
    }
    // Never substitute a weakly related folder: wrong music is worse than a
    // missing-link message that an admin can repair.
    if (best && best.score >= 0.65) return best.url;
  } catch (_) {
    // A missing search permission should not prevent rows with stored links from working.
  }
  return '';
}
