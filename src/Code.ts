const SONG_SHEET = 'Songs';

type Row = Record<string, unknown>;

function getSheetByName(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function readRows(sheetName: string): Row[] {
  const sh = getSheetByName(sheetName);
  const vals = sh.getDataRange().getValues();
  if (!vals.length) return [];
  const headers = (vals.shift() as unknown[]).map(h => String(h ?? '').trim());
  return vals.map(r => {
    const o: Row = {};
    headers.forEach((h, i) => (o[h] = r[i]));
    return o;
  });
}

function getSongs(): Row[] {
  return readRows(SONG_SHEET);
}

function doGet(e?: GoogleAppsScript.Events.DoGet) {
  // JSON for fetch() requests
  if (e?.parameter?.json === '1') {
    const rows = getSongs();
    return ContentService
      .createTextOutput(JSON.stringify({ rows }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // HTML app
  return HtmlService.createHtmlOutputFromFile('html/index')
    .setTitle('Worship Planner');
}
