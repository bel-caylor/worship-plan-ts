// src/features/services.ts
import { SERVICES_SHEET, PLANNER_SHEET, SERVICES_COL, ORDER_SHEET, ORDER_COL } from '../constants';
import { getSheetByName } from '../util/sheets';

export type AddServiceInput = {
  date?: string;      // e.g. '2025-06-01'
  time?: string;      // e.g. '10:00 AM'
  type?: string;      // ServiceType
  leader?: string;
  preacher?: string;
  scripture?: string;
  scriptureText?: string; // optional override text
  // optional free text fields
  theme?: string;
  notes?: string;
};

// --- Normalization helpers ---
function normalizeDisplayName(s: string): string {
  const clean = String(s ?? '')
    .trim()
    .replace(/\s+/g, ' ');
  if (!clean) return '';
  return clean
    .split(' ')
    .map(w => (w ? w[0].toUpperCase() + w.slice(1).toLowerCase() : w))
    .join(' ');
}

export function addService(input: AddServiceInput) {
  const sh = getSheetByName(SERVICES_SHEET);

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const idIdx = col(SERVICES_COL.id);
  const dateIdx = col(SERVICES_COL.date);
  const timeIdx = col(SERVICES_COL.time);
  const typeIdx = col(SERVICES_COL.type);
  const leaderIdx = col(SERVICES_COL.leader);
  const preacherIdx = col(SERVICES_COL.preacher);
  const scriptureIdx = col(SERVICES_COL.scripture);
  const scriptureTextIdx = (() => {
    const i1 = col(SERVICES_COL.scriptureText);
    if (i1 >= 0) return i1;
    const i2 = col('ScriptureText');
    return i2 >= 0 ? i2 : -1;
  })();
  const themeIdx = col(SERVICES_COL.theme);
  const keywordsIdx = col(SERVICES_COL.keywords);
  const notesIdx = col(SERVICES_COL.notes);

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

  // Before writing, check for duplicate ServiceID if we can compute one
  if (computedId && idIdx >= 0) {
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const idColA1 = sh.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().map(r => String(r[0] ?? '').trim());
      const exists = idColA1.some(v => v === computedId);
      if (exists) {
        throw new Error(`Service already exists: ${computedId}`);
      }
    }
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
  if (leaderIdx >= 0) vals[leaderIdx] = normalizeDisplayName(input.leader ?? '');
  if (preacherIdx >= 0) vals[preacherIdx] = normalizeDisplayName(input.preacher ?? '');
  if (scriptureIdx >= 0) vals[scriptureIdx] = input.scripture ?? '';
  // Populate scripture text: prefer explicit override; otherwise fetch via API when reference provided
  try {
    if (scriptureTextIdx >= 0) {
      const override = String((input as any).scriptureText || '').trim();
      if (override) {
        vals[scriptureTextIdx] = override;
      } else if (String(input.scripture || '').trim()) {
        const { text } = esvPassage({ reference: String(input.scripture) });
        vals[scriptureTextIdx] = text || '';
      }
    }
  } catch (_) {
    // ignore fetch failures; leave cell blank
  }
  if (themeIdx >= 0) vals[themeIdx] = input.theme ?? '';
  if (keywordsIdx >= 0) (vals as any)[keywordsIdx] = (input as any).keywords ?? '';
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

export function listServices() {
  const sh = getSheetByName(SERVICES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { items: [] };

  // Try cached response keyed by sheet shape (lastRow/lastCol)
  try {
    const ver = `${lastRow}-${lastCol}`;
    const cache = CacheService.getDocumentCache();
    const cached = cache.get('listServices:v1');
    if (cached) {
      const obj = JSON.parse(cached);
      if (obj && obj.ver === ver && Array.isArray(obj.items)) {
        return { items: obj.items };
      }
    }
    const result = (() => {
      // fallthrough to compute fresh
      const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
      const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
      const idIdx = col(SERVICES_COL.id);
      const dateIdx = col(SERVICES_COL.date);
      const timeIdx = col(SERVICES_COL.time);
      const typeIdx = col(SERVICES_COL.type);
      const leaderIdx = col(SERVICES_COL.leader);
      const preacherIdx = col(SERVICES_COL.preacher);
      const scriptureIdx = col(SERVICES_COL.scripture);
      const scriptureTextIdx = (() => {
        const i1 = col(SERVICES_COL.scriptureText);
        if (i1 >= 0) return i1;
        const i2 = col('ScriptureText');
        return i2 >= 0 ? i2 : -1;
      })();
      const themeIdx = col(SERVICES_COL.theme);
      const notesIdx = col(SERVICES_COL.notes);

      const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      const toISO = (v: any) => {
        try {
          if (v instanceof Date && !isNaN(v.getTime())) {
            const y = v.getFullYear();
            const m = String(v.getMonth() + 1).padStart(2, '0');
            const d = String(v.getDate()).padStart(2, '0');
            return `${y}-${m}-${d}`;
          }
          const s = String(v ?? '').trim();
          if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
          return s;
        } catch (_) {
          return String(v ?? '')
        }
      };
      const toTime = (v: any) => {
        try {
          if (v instanceof Date && !isNaN(v.getTime())) {
            const tz = Session.getScriptTimeZone?.() || 'Etc/UTC';
            return Utilities.formatDate(v, tz as string, 'h:mm a');
          }
          const s = String(v ?? '').trim();
          // If it's already a friendly time string, keep it
          if (!s) return '';
          const m = s.match(/^(\d{1,2})(?::(\d{2}))(?:\s*:(\d{2}))?\s*(AM|PM)?$/i);
          if (m) {
            const mm = m[2] || '00';
            const ap = (m[4] || '').toUpperCase();
            const hh = m[1];
            return `${hh}:${mm}${ap ? ' ' + ap : ''}`.trim();
          }
          return s;
        } catch(_) {
          return String(v ?? '');
        }
      };

      const items = body.map(r => ({
        id: idIdx >= 0 ? String(r[idIdx] ?? '') : '',
        date: dateIdx >= 0 ? toISO(r[dateIdx]) : '',
        time: timeIdx >= 0 ? toTime(r[timeIdx]) : '',
        type: typeIdx >= 0 ? String(r[typeIdx] ?? '') : '',
        leader: leaderIdx >= 0 ? String(r[leaderIdx] ?? '') : '',
        preacher: preacherIdx >= 0 ? String(r[preacherIdx] ?? '') : '',
        scripture: scriptureIdx >= 0 ? String(r[scriptureIdx] ?? '') : '',
        scriptureText: scriptureTextIdx >= 0 ? String(r[scriptureTextIdx] ?? '') : '',
        theme: themeIdx >= 0 ? String(r[themeIdx] ?? '') : '',
        notes: notesIdx >= 0 ? String(r[notesIdx] ?? '') : ''
      }));

      const toKey = (it: any) => (it.id && String(it.id)) || `${it.date || ''} ${it.time || ''}`;
      items.sort((a, b) => String(toKey(b)).localeCompare(String(toKey(a))));
      return { items };
    })();
    try { CacheService.getDocumentCache().put('listServices:v1', JSON.stringify({ ver, items: result.items }), 300); } catch(_) {}
    return result;
  } catch (_) { /* ignore cache errors */ }

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const idIdx = col(SERVICES_COL.id);
  const dateIdx = col(SERVICES_COL.date);
  const timeIdx = col(SERVICES_COL.time);
  const typeIdx = col(SERVICES_COL.type);
  const leaderIdx = col(SERVICES_COL.leader);
  const preacherIdx = col(SERVICES_COL.preacher);
  const scriptureIdx = col(SERVICES_COL.scripture);
  const scriptureTextIdx = (() => {
    const i1 = col(SERVICES_COL.scriptureText);
    if (i1 >= 0) return i1;
    const i2 = col('ScriptureText');
    return i2 >= 0 ? i2 : -1;
  })();
  const themeIdx = col(SERVICES_COL.theme);
  const keywordsIdx = col(SERVICES_COL.keywords);
  const notesIdx = col(SERVICES_COL.notes);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const toISO = (v: any) => {
    try {
      if (v instanceof Date && !isNaN(v.getTime())) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, '0');
        const d = String(v.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
      const s = String(v ?? '').trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
      return s;
    } catch (_) {
      return String(v ?? '')
    }
  };
  const toTime = (v: any) => {
    try {
      if (v instanceof Date && !isNaN(v.getTime())) {
        const tz = Session.getScriptTimeZone?.() || 'Etc/UTC';
        return Utilities.formatDate(v, tz as string, 'h:mm a');
      }
      const s = String(v ?? '').trim();
      // If it's already a friendly time string, keep it
      if (!s) return '';
      // Handle cases like 12:00:00 AM -> 12:00 AM
      const m = s.match(/^(\d{1,2})(?::(\d{2}))(?:\s*:(\d{2}))?\s*(AM|PM)?$/i);
      if (m) {
        const mm = m[2] || '00';
        const ap = (m[4] || '').toUpperCase();
        const hh = m[1];
        return `${hh}:${mm}${ap ? ' ' + ap : ''}`.trim();
      }
      return s;
    } catch(_) {
      return String(v ?? '');
    }
  };

  const items = body.map(r => ({
    id: idIdx >= 0 ? String(r[idIdx] ?? '') : '',
    date: dateIdx >= 0 ? toISO(r[dateIdx]) : '',
    time: timeIdx >= 0 ? toTime(r[timeIdx]) : '',
    type: typeIdx >= 0 ? String(r[typeIdx] ?? '') : '',
    leader: leaderIdx >= 0 ? String(r[leaderIdx] ?? '') : '',
    preacher: preacherIdx >= 0 ? String(r[preacherIdx] ?? '') : '',
    scripture: scriptureIdx >= 0 ? String(r[scriptureIdx] ?? '') : '',
    scriptureText: scriptureTextIdx >= 0 ? String(r[scriptureTextIdx] ?? '') : '',
    theme: themeIdx >= 0 ? String(r[themeIdx] ?? '') : '',
    keywords: keywordsIdx >= 0 ? String(r[keywordsIdx] ?? '') : '',
    notes: notesIdx >= 0 ? String(r[notesIdx] ?? '') : ''
  }));

  // Sort descending by ServiceID (fallback to date+time)
  const toKey = (it: any) => (it.id && String(it.id)) || `${it.date || ''} ${it.time || ''}`;
  items.sort((a, b) => String(toKey(b)).localeCompare(String(toKey(a))));

  return { items };
}

export function saveService(input: AddServiceInput & { id?: string }) {
  const sh = getSheetByName(SERVICES_SHEET);

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const idIdx = col(SERVICES_COL.id);
  const dateIdx = col(SERVICES_COL.date);
  const timeIdx = col(SERVICES_COL.time);
  const typeIdx = col(SERVICES_COL.type);
  const leaderIdx = col(SERVICES_COL.leader);
  const preacherIdx = col(SERVICES_COL.preacher);
  const scriptureIdx = col(SERVICES_COL.scripture);
  const scriptureTextIdx = (() => {
    const i1 = col(SERVICES_COL.scriptureText);
    if (i1 >= 0) return i1;
    const i2 = col('ScriptureText');
    return i2 >= 0 ? i2 : -1;
  })();
  const themeIdx = col(SERVICES_COL.theme);
  const keywordsIdx = col(SERVICES_COL.keywords);
  const notesIdx = col(SERVICES_COL.notes);

  // Compute an ID from provided date/time just like addService
  let computedId = '';
  try {
    let y = 0, m = 0, d = 0;
    const inDate = input.date;
    if (inDate instanceof Date && !isNaN(inDate.getTime())) {
      y = inDate.getFullYear(); m = inDate.getMonth() + 1; d = inDate.getDate();
    } else if (typeof inDate === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(inDate)) {
      const [yy, mm, dd] = inDate.split('-').map(Number);
      y = yy; m = mm; d = dd;
    }

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
      hh = 10; min = 0;
    }

    if (y && m && d) {
      const MM = String(m).padStart(2, '0');
      const DD = String(d).padStart(2, '0');
      let h12 = hh % 12; if (h12 === 0) h12 = 12;
      const suffix = hh < 12 ? 'am' : 'pm';
      const minutePart = min ? `:${String(min).padStart(2, '0')}` : '';
      computedId = `${y}-${MM}-${DD}_${h12}${minutePart}${suffix}`;
    }
  } catch (_) {
    computedId = '';
  }

  const originalId = String(input.id || '').trim();
  const newId = computedId || originalId;

  // Build row data according to headers
  const vals: any[] = Array.from({ length: lastCol }, () => '');
  if (idIdx >= 0) vals[idIdx] = newId;
  if (dateIdx >= 0) {
    const d = String(input.date || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(d)) {
      const [y, m, day] = d.split('-').map(Number);
      vals[dateIdx] = new Date(y, (m - 1), day);
    } else {
      vals[dateIdx] = d;
    }
  }
  if (timeIdx >= 0) vals[timeIdx] = input.time ?? '';
  if (typeIdx >= 0) vals[typeIdx] = input.type ?? '';
  if (leaderIdx >= 0) vals[leaderIdx] = normalizeDisplayName(input.leader ?? '');
  if (preacherIdx >= 0) vals[preacherIdx] = normalizeDisplayName(input.preacher ?? '');
  if (scriptureIdx >= 0) vals[scriptureIdx] = input.scripture ?? '';
  try {
    if (scriptureTextIdx >= 0) {
      const override = String((input as any).scriptureText || '').trim();
      if (override) {
        vals[scriptureTextIdx] = override;
      } else if (String(input.scripture || '').trim()) {
        const { text } = esvPassage({ reference: String(input.scripture) });
        vals[scriptureTextIdx] = text || '';
      }
    }
  } catch (_) { /* ignore */ }
  if (themeIdx >= 0) vals[themeIdx] = input.theme ?? '';
  if (keywordsIdx >= 0) (vals as any)[keywordsIdx] = (input as any).keywords ?? '';
  if (notesIdx >= 0) vals[notesIdx] = input.notes ?? '';

  // Find row by originalId (preferred) or by computedId
  const lastRow = sh.getLastRow();
  let rowIdx = -1; // 0-based into data region; will convert to absolute later
  if (idIdx >= 0 && lastRow >= 2) {
    const idVals = sh.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().map(r => String(r[0] ?? '').trim());
    if (originalId) {
      rowIdx = idVals.findIndex(v => v === originalId);
    }
    if (rowIdx < 0 && newId) {
      rowIdx = idVals.findIndex(v => v === newId);
    }

    // Duplicate check when changing ID
    if (originalId && newId && newId !== originalId) {
      const dup = idVals.some(v => v === newId);
      if (dup) throw new Error(`Service already exists: ${newId}`);
    }
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    if (rowIdx >= 0) {
      // Update the existing row (rowIdx maps to sheet row = 2 + rowIdx)
      sh.getRange(2 + rowIdx, 1, 1, lastCol).setValues([vals]);
      return { id: newId };
    } else {
      // Fallback to add if not found
      sh.appendRow(vals);
      return { id: newId };
    }
  } finally {
    lock.releaseLock();
  }
}

export function deleteService(input: { id?: string } | string) {
  const id = typeof input === 'string' ? input : String((input as any)?.id || '').trim();
  const serviceId = String(id || '').trim();
  if (!serviceId) throw new Error('id required');

  // Delete row from Services and any related rows from Order
  // Services
  const sh = getSheetByName(SERVICES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const idIdx = col(SERVICES_COL.id);
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    if (idIdx >= 0 && lastRow >= 2) {
      const ids = sh.getRange(2, idIdx + 1, lastRow - 1, 1).getValues().map(r => String(r[0] ?? '').trim());
      for (let i = ids.length - 1; i >= 0; i--) {
        if (ids[i] === serviceId) sh.deleteRow(2 + i);
      }
    }
  } finally {
    lock.releaseLock();
  }

  // Related order rows
  try {
    const oh = getSheetByName(ORDER_SHEET);
    const oLastRow = oh.getLastRow();
    const oLastCol = oh.getLastColumn();
    const oHeaders = oh.getRange(1, 1, 1, oLastCol).getValues()[0].map(v => String(v ?? '').trim());
    const oCol = (name: string) => oHeaders.findIndex(h => h.toLowerCase() === name.toLowerCase());
    const serviceIdx = oCol(ORDER_COL.serviceId);
    if (serviceIdx >= 0 && oLastRow >= 2) {
      const ids = oh.getRange(2, serviceIdx + 1, oLastRow - 1, 1).getValues().map(r => String(r[0] ?? '').trim());
      for (let i = ids.length - 1; i >= 0; i--) if (ids[i] === serviceId) oh.deleteRow(2 + i);
    }
  } catch (_) { /* ignore */ }

  return { ok: true };
}

export function getServicePeople() {
  const toDisplay = (s: string) => s
    .trim()
    .replace(/\s+/g, ' ')
    .split(' ')
    .map(w => (w ? w[0].toUpperCase() + w.slice(1).toLowerCase() : w))
    .join(' ');
  const toKey = (s: string) => s.trim().replace(/\s+/g, ' ').toLowerCase();

  const merge = (map: Map<string, string>, vals: any[]) => {
    for (const v of vals) {
      const raw = String(v ?? '');
      const key = toKey(raw);
      if (!key) continue;
      if (!map.has(key)) map.set(key, toDisplay(raw));
    }
  };

  const leaders = new Map<string, string>();
  const preachers = new Map<string, string>();

  // From Services sheet
  try {
    const sh = getSheetByName(SERVICES_SHEET);
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow >= 2 && lastCol >= 1) {
      const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
      const normIdx = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
      const leaderIdx = normIdx(SERVICES_COL.leader);
      const preacherIdx = normIdx(SERVICES_COL.preacher);
      const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      if (leaderIdx >= 0) merge(leaders, body.map(r => r[leaderIdx]));
      if (preacherIdx >= 0) merge(preachers, body.map(r => r[preacherIdx]));
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
      const leaderIdx = normIdx(SERVICES_COL.leader);
      const preacherIdx = normIdx(SERVICES_COL.preacher);
      const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      if (leaderIdx >= 0) merge(leaders, body.map(r => r[leaderIdx]));
      if (preacherIdx >= 0) merge(preachers, body.map(r => r[preacherIdx]));
    }
  } catch (_) { /* ignore */ }

  const sort = (a: string, b: string) => a.localeCompare(b);
  return {
    leaders: Array.from(leaders.values()).sort(sort),
    preachers: Array.from(preachers.values()).sort(sort)
  };
}

export function esvPassage(input: { reference: string, html?: boolean }) {
  const reference = String(input?.reference || '').trim();
  if (!reference) return { reference, text: '' };

  const props = PropertiesService.getScriptProperties();
  const token = String(props.getProperty('ESV_API_TOKEN') || '');
  if (!token) {
    return { reference, text: '', html: '', error: 'ESV_API_TOKEN not set in Script Properties' };
  }

  const url = 'https://api.esv.org/v3/passage/text/?' +
    'q=' + encodeURIComponent(reference) +
    '&include-passage-references=false' + // don't echo the reference header
    '&include-footnotes=false' +
    '&include-headings=false' +
    '&include-short-copyright=false' +
    '&include-verse-numbers=false' +
    '&indent-poetry=false' +
    '&indent-using=spaces' +
    '&indent-paragraphs=0';

  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Token ' + token } });
  const data = JSON.parse(res.getContentText());
  let text = (data && data.passages && data.passages[0]) ? String(data.passages[0]) : '';
  // Remove bracketed footnote remnants or stray markers just in case
  text = text.replace(/\s*\[\d+\]\s*/g, ' ');
  // Normalize whitespace: trim, collapse 3+ newlines to 2, normalize CRLF
  text = text.replace(/\r\n?/g, '\n');
  text = text.replace(/\n{3,}/g, '\n\n').trim();
  let html = '';
  try {
    const hurl = 'https://api.esv.org/v3/passage/html/?' +
      'q=' + encodeURIComponent(reference) +
      '&include-passage-references=false' +
      '&include-footnotes=false' +
      '&include-headings=false' +
      '&include-short-copyright=false' +
      '&include-verse-numbers=true' +
      '&inline-styles=false';
    const hres = UrlFetchApp.fetch(hurl, { headers: { Authorization: 'Token ' + token } });
    const hdata = JSON.parse(hres.getContentText());
    html = (hdata && hdata.passages && hdata.passages[0]) ? String(hdata.passages[0]) : '';
    // Basic cleanup: remove outer wrappers if present
    html = html.replace(/<p class=".*?">/g, '<p>').replace(/<h\d[^>]*>.*?<\/h\d>/g, '');
  } catch (_) { /* ignore html errors */ }
  return { reference, text, html };
}
