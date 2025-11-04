// src/util/sheets.ts
import { SONG_SHEET, SONG_COL_NAME } from '../constants';

export function readRows(sheetName: string): Row[] {
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

export function getSongs(): Row[] {
    return readRows(SONG_SHEET);
}

export function getSheetByName(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

export function getHeaders(sh: GoogleAppsScript.Spreadsheet.Sheet) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const colMap: Record<string, number> = {};
  headers.forEach((h, i) => (colMap[h] = i));
  if (!(SONG_COL_NAME in colMap)) throw new Error(`Header "${SONG_COL_NAME}" not found`);
  return { headers, colMap };
}

export function ensureColumn(
    sh: GoogleAppsScript.Spreadsheet.Sheet,
    headers: string[],
    colMap: Record<string, number>,
    name: string
) {
    if (name in colMap) return;
    sh.insertColumnAfter(headers.length);
    const newCol = headers.length + 1;
    sh.getRange(1, newCol).setValue(name);
    headers.push(name);
    colMap[name] = headers.length - 1;
}

export function findHeaderIndex(headers: string[], candidates: string[]): number {
    const H = headers.map(norm);
    for (const cand of candidates) {
        const i = H.indexOf(norm(cand));
        if (i >= 0) return i;
    }
    return -1;
}

export function findManyHeaderIndices(headers: string[], labels: string[]): number[] {
    const H = headers.map(norm);
    const idxs: number[] = [];
    for (const label of labels) {
        const i = H.indexOf(norm(label));
        if (i >= 0) idxs.push(i);
    }
    return idxs;
}
