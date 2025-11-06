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

export type ListServicesOptions = {
  startDate?: string;
  endDate?: string;
  includePast?: boolean;
  limit?: number;
  sort?: 'asc' | 'desc';
};

export type CreateServicesBatchInput = {
  startDate?: string;
  weeks?: number;
};

export type ServiceItem = {
  id: string;
  date: string;
  time: string;
  type: string;
  leader: string;
  preacher: string;
  scripture: string;
  scriptureText: string;
  theme: string;
  keywords: string;
  notes: string;
};

const SERVICES_CACHE_KEY = 'listServices:v1';
const DEFAULT_SERVICE_TIME = '10:00 AM';
const DEFAULT_LEADER = 'Darden';
const DEFAULT_PREACHER = 'Tom';
const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

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

function toSheetDateValue(input: any): Date | string {
  try {
    const safeDate = (y: number, m: number, d: number) => new Date(Date.UTC(y, m, d, 12, 0, 0));
    if (input instanceof Date && !isNaN(input.getTime())) {
      return safeDate(input.getFullYear(), input.getMonth(), input.getDate());
    }
    const s = String(input ?? '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      const [yy, mm, dd] = s.split('-').map(Number);
      return safeDate(yy, mm - 1, dd);
    }
    return s;
  } catch (_) {
    return String(input ?? '');
  }
}

function deriveKeywords(text: any): string {
  const s = String(text || '').toLowerCase();
  if (!s) return '';
  const tokens = s.replace(/[^a-z\s']/g, ' ').split(/\s+/).map(t => t.replace(/^'+|'+$/g, '')).filter(Boolean);
  if (!tokens.length) return '';
  const stop = new Set([
    'the','and','of','to','in','that','it','is','for','on','with','as','at','by','be','he','she','they','we','you','i','a','an','from','this','these','those','are','was','were','his','her','their','our','your','but','not','so','or','if','then','there','here','who','whom','which','what','when','where','why','how','have','has','had','do','did','does','will','would','shall','should','can','could','may','might','let','us',
    'him','them','me','my','mine','yours','ours','hers','theirs','whoever','whosoever','whomever','whose','into','unto','onto','upon','within','without','among','between','before','after','above','below','over','under','again','also','all','any','each','every','some','no','nor','one','thing','things','because'
  ]);
  const lemma = (w: string): string => {
    if (!w) return '';
    if (/^bright(?:ness)?$/.test(w) || /^shine(?:s|r|rs|d|ing)?$/.test(w)) return 'light';
    if (/^light(?:s|er|est|ness)?$/.test(w)) return 'light';
    if (/^dark(?:ness|er|est|s)?$/.test(w)) return 'darkness';
    if (/^judg(?:e|es|ed|ing|ment|ments)$/.test(w) || /^condemn(?:ed|s|ing|ation|ations)?$/.test(w)) return 'judgment';
    if (/^believ(?:e|es|ed|ing|er|ers)?$/.test(w)) return 'believe';
    if (/^baptiz(?:e|es|ed|ing)?$/.test(w) || /^baptism(?:s)?$/.test(w) || /^baptist(?:s)?$/.test(w)) return 'baptism';
    if (/^come(?:s|r|rs|ing)?$/.test(w) || w === 'came') return 'come';
    if (w.length > 4 && /s$/.test(w)) return w.replace(/s$/, '');
    return w;
  };
  const counts = new Map<string, number>();
  for (const t of tokens) {
    if (stop.has(t) || t.length < 3) continue;
    const k = lemma(t);
    if (!k || stop.has(k) || k.length < 3) continue;
    counts.set(k, (counts.get(k) || 0) + 1);
  }
  const top = Array.from(counts.entries()).sort((a, b) => b[1] - a[1]).slice(0, 12).map(([k]) => k);
  const pretty = (w: string) => w.replace(/^\w/, c => c.toUpperCase());
  return top.map(pretty).join(', ');
}

const isoFromDate = (date: Date) => {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

const dateFromISO = (iso: string): Date | null => {
  if (!ISO_DATE_RE.test(String(iso || ''))) return null;
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d);
};

const normalizeIso = (value?: string | Date): string | null => {
  if (!value && value !== '') return null;
  if (value instanceof Date && !isNaN(value.getTime())) return isoFromDate(value);
  const s = String(value ?? '').trim();
  return ISO_DATE_RE.test(s) ? s : null;
};

const nextSundayOnOrAfter = (date: Date): Date => {
  const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const delta = (7 - copy.getDay()) % 7;
  if (delta) copy.setDate(copy.getDate() + delta);
  return copy;
};

const defaultServiceTypeForDate = (date: Date): string => {
  const nth = Math.floor((date.getDate() - 1) / 7) + 1;
  return (nth === 1 || nth === 3 || nth === 5) ? 'Communion' : 'Offering';
};

const todayISO = () => {
  const tz = (Session.getScriptTimeZone && Session.getScriptTimeZone()) || 'Etc/UTC';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
};

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
  if (keywordsIdx >= 0) {
    const provided = String((input as any).keywords ?? '').trim();
    const textSource = provided
      ? ''
      : (scriptureTextIdx >= 0 ? String(vals[scriptureTextIdx] ?? '') : String((input as any).scriptureText ?? ''));
    const keywords = provided || deriveKeywords(textSource);
    (vals as any)[keywordsIdx] = keywords;
  }
  if (notesIdx >= 0) vals[notesIdx] = input.notes ?? '';

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    sh.appendRow(vals);
  } finally {
    lock.releaseLock();
  }

  try { CacheService.getDocumentCache().remove(SERVICES_CACHE_KEY); } catch (_) {}
  return { id: computedId };
}

function fetchServicesUnfiltered(): ServiceItem[] {
  const sh = getSheetByName(SERVICES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  // Try cached response keyed by sheet shape (lastRow/lastCol)
  try {
    const updatedAt = (() => {
      try { return SpreadsheetApp.getActive().getLastUpdated()?.getTime() || 0; } catch (_) { return 0; }
    })();
    const ver = `${lastRow}-${lastCol}-${updatedAt}`;
    const cache = CacheService.getDocumentCache();
    const cached = cache.get(SERVICES_CACHE_KEY);
    if (cached) {
      const obj = JSON.parse(cached);
      if (obj && obj.ver === ver && Array.isArray(obj.items)) {
        return obj.items as ServiceItem[];
      }
    }
    const items = (() => {
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

      const rows = body.map(r => ({
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

      const toKey = (it: any) => (it.id && String(it.id)) || `${it.date || ''} ${it.time || ''}`;
      rows.sort((a, b) => String(toKey(b)).localeCompare(String(toKey(a))));
      return rows as ServiceItem[];
    })();
    try { CacheService.getDocumentCache().put(SERVICES_CACHE_KEY, JSON.stringify({ ver, items }), 300); } catch(_) {}
    return items;
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

  return items as ServiceItem[];
}

const serviceSortKey = (item: ServiceItem) => (item.id && String(item.id)) || `${item.date || ''} ${item.time || ''}`.trim();

function applyServiceFilters(items: ServiceItem[], opts?: ListServicesOptions): ServiceItem[] {
  let result = Array.isArray(items) ? items.slice() : [];
  if (!opts) return result;
  const start = normalizeIso(opts.startDate || undefined);
  if (start) {
    result = result.filter(item => !item.date || item.date >= start);
  }
  const end = normalizeIso(opts.endDate || undefined);
  if (end) {
    result = result.filter(item => !item.date || item.date <= end);
  }
  if (opts.includePast === false) {
    const cutoff = todayISO();
    result = result.filter(item => !item.date || item.date >= cutoff);
  }
  if (opts.sort === 'asc') {
    result.sort((a, b) => serviceSortKey(a).localeCompare(serviceSortKey(b)));
  } else if (opts.sort === 'desc') {
    result.sort((a, b) => serviceSortKey(b).localeCompare(serviceSortKey(a)));
  }
  const limit = typeof opts.limit === 'number' ? Math.max(0, Math.floor(opts.limit)) : 0;
  if (limit > 0 && result.length > limit) {
    result = result.slice(0, limit);
  }
  return result;
}

export function listServices(opts?: ListServicesOptions) {
  const all = fetchServicesUnfiltered();
  return { items: applyServiceFilters(all, opts) };
}

export function createServicesBatch(input?: CreateServicesBatchInput) {
  const weeksValue = Number(input?.weeks);
  const weeksRaw = Number.isFinite(weeksValue) ? Math.floor(weeksValue) : NaN;
  const weeks = Math.min(52, Math.max(1, isNaN(weeksRaw) ? 12 : weeksRaw));
  const startIso = normalizeIso(input?.startDate || '') || isoFromDate(nextSundayOnOrAfter(new Date()));
  const startDate = dateFromISO(startIso) || nextSundayOnOrAfter(new Date());
  const firstSunday = nextSundayOnOrAfter(startDate);
  const schedule: { iso: string; date: Date }[] = [];
  for (let i = 0; i < weeks; i++) {
    const iter = new Date(firstSunday.getFullYear(), firstSunday.getMonth(), firstSunday.getDate() + (i * 7));
    schedule.push({ iso: isoFromDate(iter), date: iter });
  }

  const sh = getSheetByName(SERVICES_SHEET);
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error(`Sheet ${SERVICES_SHEET} is missing headers`);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const idIdx = col(SERVICES_COL.id);
  if (idIdx === -1) throw new Error(`Column "${SERVICES_COL.id}" not found in ${SERVICES_SHEET}`);
  const dateIdx = col(SERVICES_COL.date);
  const timeIdx = col(SERVICES_COL.time);
  const typeIdx = col(SERVICES_COL.type);
  const leaderIdx = col(SERVICES_COL.leader);
  const preacherIdx = col(SERVICES_COL.preacher);

  const created: { id: string; date: string; time: string; type: string }[] = [];
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const existing = new Set<string>();
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const ids = sh.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
      ids.forEach(row => {
        const id = String((row && row[0]) ?? '').trim();
        if (id) existing.add(id);
      });
    }
    const rows: any[][] = [];
    for (const entry of schedule) {
      const serviceId = `${entry.iso}_10am`;
      if (existing.has(serviceId)) continue;
      existing.add(serviceId);
      const row = Array.from({ length: lastCol }, () => '');
      row[idIdx] = serviceId;
      if (dateIdx >= 0) row[dateIdx] = new Date(entry.date.getFullYear(), entry.date.getMonth(), entry.date.getDate());
      if (timeIdx >= 0) row[timeIdx] = DEFAULT_SERVICE_TIME;
      const svcType = defaultServiceTypeForDate(entry.date);
      if (typeIdx >= 0) row[typeIdx] = svcType;
      if (leaderIdx >= 0) row[leaderIdx] = DEFAULT_LEADER;
      if (preacherIdx >= 0) row[preacherIdx] = DEFAULT_PREACHER;
      rows.push(row);
      created.push({ id: serviceId, date: entry.iso, time: DEFAULT_SERVICE_TIME, type: svcType });
    }
    if (rows.length) {
      const startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, rows.length, lastCol).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }
  try { CacheService.getDocumentCache().remove(SERVICES_CACHE_KEY); } catch (_) {}
  return { created };
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
  if (dateIdx >= 0) vals[dateIdx] = toSheetDateValue(input.date);
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
  if (keywordsIdx >= 0) {
    const provided = String((input as any).keywords ?? '').trim();
    const textSource = provided
      ? ''
      : (scriptureTextIdx >= 0 ? String(vals[scriptureTextIdx] ?? '') : String((input as any).scriptureText ?? ''));
    const keywords = provided || deriveKeywords(textSource);
    (vals as any)[keywordsIdx] = keywords;
  }
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
  let resultId = newId;
  try {
    if (rowIdx >= 0) {
      // Update the existing row (rowIdx maps to sheet row = 2 + rowIdx)
      sh.getRange(2 + rowIdx, 1, 1, lastCol).setValues([vals]);
    } else {
      // Fallback to add if not found
      sh.appendRow(vals);
    }
  } finally {
    lock.releaseLock();
  }
  try { CacheService.getDocumentCache().remove(SERVICES_CACHE_KEY); } catch (_) {}
  return { id: resultId };
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

  try { CacheService.getDocumentCache().remove(SERVICES_CACHE_KEY); } catch (_) {}
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
