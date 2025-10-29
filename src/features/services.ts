// src/features/services.ts
import { SERVICES_SHEET, PLANNER_SHEET } from '../constants';
import { getSheetByName } from '../util/sheets';

export type AddServiceInput = {
  date?: string;      // e.g. '2025-06-01'
  time?: string;      // e.g. '10:00 AM'
  type?: string;      // ServiceType
  leader?: string;
  sermon?: string;
  scripture?: string;
  // optional free text fields
  theme?: string;
  notes?: string;
};

export function addService(input: AddServiceInput) {
  const sh = getSheetByName(SERVICES_SHEET);

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const idIdx = col('ServiceID');
  const dateIdx = col('Date');
  const timeIdx = col('Time');
  const typeIdx = col('ServiceType');
  const leaderIdx = col('Leader');
  const sermonIdx = col('Sermon');
  const scriptureIdx = col('Scripture');
  const scriptureTextIdx = (() => {
    const i1 = col('Scripture Text');
    if (i1 >= 0) return i1;
    const i2 = col('ScriptureText');
    return i2 >= 0 ? i2 : -1;
  })();
  const themeIdx = col('Theme');
  const notesIdx = col('Notes');

  // Build a deterministic ServiceID from date + time, e.g., 2025-11-02_10am
  let computedId = '';
  try {
    // Extract date parts
    let y = 0, m = 0, d = 0;
    const inDate = input.date;
    if (inDate instanceof Date && !isNaN(inDate.getTime())) {
      y = inDate.getFullYear(); m = inDate.getMonth() + 1; d = inDate.getDate();
    } else if (typeof inDate === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(inDate)) {
      const [yy, mm, dd] = inDate.split('-').map(Number);
      y = yy; m = mm; d = dd;
    }

    // Extract time parts (support 'h:mm AM/PM', 'h AM/PM', 'HH:mm')
    let hh = 0, min = 0;
    const t = String(input.time || '').trim();
    if (t) {
      const ampm = t.match(/\b(AM|PM)\b/i)?.[1]?.toUpperCase() || '';
      const nums = t.match(/(\d{1,2})(?::(\d{2}))?/);
      if (nums) {
        hh = Number(nums[1]);
        min = nums[2] ? Number(nums[2]) : 0;
        if (ampm === 'AM') {
          if (hh === 12) hh = 0;
        } else if (ampm === 'PM') {
          if (hh !== 12) hh += 12;
        }
      }
    } else {
      hh = 10; min = 0; // default to 10:00 if unspecified
    }

    if (y && m && d) {
      const MM = String(m).padStart(2, '0');
      const DD = String(d).padStart(2, '0');
      // Convert to 12-hour for the ID and lowercase am/pm
      let h12 = hh % 12; if (h12 === 0) h12 = 12;
      const suffix = hh < 12 ? 'am' : 'pm';
      const minutePart = min ? `:${String(min).padStart(2, '0')}` : '';
      computedId = `${y}-${MM}-${DD}_${h12}${minutePart}${suffix}`;
    }
  } catch (_) {
    computedId = '';
  }

  // Build the row sized to current header count
  const vals: any[] = Array.from({ length: lastCol }, () => '');

  if (idIdx >= 0) vals[idIdx] = computedId;

  if (dateIdx >= 0) {
    const d = String(input.date || '').trim();
    // If looks like YYYY-MM-DD, convert to Date so Sheets stores a date
    if (/^\d{4}-\d{2}-\d{2}$/.test(d)) {
      const [y, m, day] = d.split('-').map(Number);
      vals[dateIdx] = new Date(y, (m - 1), day);
    } else {
      vals[dateIdx] = d;
    }
  }
  if (timeIdx >= 0) vals[timeIdx] = input.time ?? '';
  if (typeIdx >= 0) vals[typeIdx] = input.type ?? '';
  if (leaderIdx >= 0) vals[leaderIdx] = input.leader ?? '';
  if (sermonIdx >= 0) vals[sermonIdx] = input.sermon ?? '';
  if (scriptureIdx >= 0) vals[scriptureIdx] = input.scripture ?? '';
  // Fetch scripture text via ESV API when a column is available and a reference provided
  try {
    if (scriptureTextIdx >= 0 && (input.scripture || '').trim()) {
      const { text } = esvPassage({ reference: String(input.scripture) });
      vals[scriptureTextIdx] = text || '';
    }
  } catch (_) {
    // ignore fetch failures; leave cell blank
  }
  if (themeIdx >= 0) vals[themeIdx] = input.theme ?? '';
  if (notesIdx >= 0) vals[notesIdx] = input.notes ?? '';

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    sh.appendRow(vals);
  } finally {
    lock.releaseLock();
  }

  return { id: computedId };
}

export function getServicePeople() {
  const merge = (set: Set<string>, vals: any[]) => {
    for (const v of vals) {
      const s = String(v ?? '').trim();
      if (s) set.add(s);
    }
  };

  const leaders = new Set<string>();
  const preachers = new Set<string>();

  // From Services sheet
  try {
    const sh = getSheetByName(SERVICES_SHEET);
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow >= 2 && lastCol >= 1) {
      const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
      const normIdx = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
      const leaderIdx = normIdx('Leader');
      const sermonIdx = normIdx('Sermon');
      const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      if (leaderIdx >= 0) merge(leaders, body.map(r => r[leaderIdx]));
      if (sermonIdx >= 0) merge(preachers, body.map(r => r[sermonIdx]));
    }
  } catch (_) { /* ignore */ }

  // From Weekly Planner sheet (if present)
  try {
    const sh = getSheetByName(PLANNER_SHEET);
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow >= 2 && lastCol >= 1) {
      const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
      const normIdx = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
      const leaderIdx = normIdx('Leader');
      const sermonIdx = normIdx('Sermon');
      const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      if (leaderIdx >= 0) merge(leaders, body.map(r => r[leaderIdx]));
      if (sermonIdx >= 0) merge(preachers, body.map(r => r[sermonIdx]));
    }
  } catch (_) { /* ignore */ }

  const sort = (a: string, b: string) => a.localeCompare(b);
  return { leaders: Array.from(leaders).sort(sort), preachers: Array.from(preachers).sort(sort) };
}

export function esvPassage(input: { reference: string }) {
  const reference = String(input?.reference || '').trim();
  if (!reference) return { reference, text: '' };

  const props = PropertiesService.getScriptProperties();
  const token = String(props.getProperty('ESV_API_TOKEN') || '');
  if (!token) {
    return { reference, text: '', error: 'ESV_API_TOKEN not set in Script Properties' };
  }

  const url = 'https://api.esv.org/v3/passage/text/?' +
    'q=' + encodeURIComponent(reference) +
    '&include-passage-references=false' + // don't echo the reference header
    '&include-footnotes=false' +
    '&include-headings=false' +
    '&include-short-copyright=false' +
    '&include-verse-numbers=true' +
    '&indent-poetry=false' +
    '&indent-using=spaces' +
    '&indent-paragraphs=0';

  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Token ' + token } });
  const data = JSON.parse(res.getContentText());
  let text = (data && data.passages && data.passages[0]) ? String(data.passages[0]) : '';
  // Normalize whitespace: trim, collapse 3+ newlines to 2, normalize CRLF
  text = text.replace(/\r\n?/g, '\n');
  text = text.replace(/\n{3,}/g, '\n\n').trim();
  return { reference, text };
}
