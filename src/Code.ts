/** Sheet names you use in your Worship planner */
const SONG_SHEET = 'Songs';

/** A typed row structure (adjust headers to your sheet) */
type SongRow = {
  Title: string;
  Key: string;
  Tempo: string;
  LastUsed: string | Date;
};

/** Utility: get active sheet (bound script → your current spreadsheet) */
function getSheetByName(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

/** Read all rows from a headered sheet into objects */
function readRows<T = Record<string, unknown>>(sheetName: string): T[] {
  const sh = getSheetByName(sheetName);
  const range = sh.getDataRange();
  const values = range.getValues(); // first row = headers
  if (values.length === 0) return [];
  const headers = (values.shift() as string[]).map(String);
  return values.map((row) => {
    const obj: Record<string, unknown> = {};
    headers.forEach((h, i) => (obj[h] = row[i]));
    return obj as T;
  });
}

/** Exposed server function (callable from HTML sidebar/web app) */
function getSongs(): SongRow[] {
  return readRows<SongRow>(SONG_SHEET);
}

/** Web app GET → returns JSON (great for local FE or App’s own client) */
function doGet(e: GoogleAppsScript.Events.DoGet) {
  const rows = getSongs();
  const body = JSON.stringify({ rows });
  return ContentService.createTextOutput(body).setMimeType(
    ContentService.MimeType.JSON
  );
}

/** Example: append a song (call via google.script.run or POST) */
function addSong(song: SongRow): string {
  const sh = getSheetByName(SONG_SHEET);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = (h: string) => (song as any)[h] ?? '';
  sh.appendRow(headers.map(map));
  return 'OK';
}
