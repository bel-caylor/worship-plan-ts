// src/features/services.ts
import { SERVICES_SHEET, PLANNER_SHEET, SERVICES_COL, ORDER_SHEET, ORDER_COL, YOUTUBE_STREAMS_SHEET, YOUTUBE_STREAMS_COL, SONG_PERFORMANCES_SHEET, SONG_PERFORMANCES_COL } from '../constants';
import { getSpreadsheetVersion, readDocumentCachedJson } from '../util/cache';
import { getSheetByName } from '../util/sheets';

export type AddServiceInput = {
  date?: string;      // e.g. '2025-06-01'
  time?: string;      // e.g. '10:00 AM'
  type?: string;      // ServiceType
  youtubeUrl?: string;
  leader?: string;
  preacher?: string;
  scripture?: string;
  scriptureText?: string; // optional override text
  // optional free text fields
  theme?: string;
  keywords?: string;
  notes?: string;
  suggestedSongs?: string;
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
  youtubeUrl: string;
  leader: string;
  preacher: string;
  scripture: string;
  scriptureText: string;
  theme: string;
  keywords: string;
  notes: string;
  suggestedSongs: string;
};

const SERVICES_CACHE_KEY = 'listServices:v1';
const SERVICE_PEOPLE_CACHE_KEY = 'servicePeople:v1';
const YOUTUBE_HELPER_CACHE_KEY = 'youtube-helper:v1';
const DEFAULT_SERVICE_TIME = '10:00 AM';
const DEFAULT_LEADER = 'Darden';
const DEFAULT_PREACHER = 'Tom';
const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const DEFAULT_YOUTUBE_CHANNEL_URL = 'https://www.youtube.com/@hopechurch7113';
const DEFAULT_YOUTUBE_STREAMS_URL = `${DEFAULT_YOUTUBE_CHANNEL_URL}/streams`;
const DEFAULT_YOUTUBE_TITLE_PREFIX = 'Hope Is Real';
const YOUTUBE_STREAMS_CURSOR_PROPERTY = 'youtube_streams_catalog_page_token_v2';
const YOUTUBE_API_KEY_PROPERTY = 'YOUTUBE_API_KEY';
const YOUTUBE_CHANNEL_HANDLE = '@hopechurch7113';
const AUTO_SERVICE_WEEKS_AHEAD = 12;

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
    // Use local noon so Sheets won't render the previous date in local time.
    const safeDate = (y: number, m: number, d: number) => new Date(y, m, d, 12, 0, 0);
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
  if (ISO_DATE_RE.test(s)) return s;
  const isoPrefix = s.match(/^(\d{4}-\d{2}-\d{2})[T\s]/);
  if (isoPrefix?.[1]) return isoPrefix[1];
  const slashMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\b|\s)/);
  if (slashMatch) {
    const month = Number(slashMatch[1]);
    const day = Number(slashMatch[2]);
    const year = Number(slashMatch[3]);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 2000) {
      return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }
  return null;
};

const nextSundayOnOrAfter = (date: Date): Date => {
  const copy = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const delta = (7 - copy.getDay()) % 7;
  if (delta) copy.setDate(copy.getDate() + delta);
  return copy;
};

const addDays = (date: Date, days: number): Date =>
  new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);

const deriveDateFromServiceId = (id: string): string => {
  const m = String(id || '').match(/^(\d{4}-\d{2}-\d{2})_/);
  return m ? m[1] : '';
};

const deriveTimeFromServiceId = (id: string): string => {
  const m = String(id || '').match(/_(\d{1,2})(?::(\d{2}))?(am|pm)\b/i);
  if (!m) return '';
  const hour = Number(m[1] || 0);
  const minutes = m[2] ? m[2].padStart(2, '0') : '00';
  const mer = (m[3] || '').toUpperCase();
  if (!hour || !mer) return '';
  return `${hour}:${minutes} ${mer}`;
};

const canonicalServiceDate = (serviceId: string, sheetValue?: unknown): string => {
  return deriveDateFromServiceId(serviceId) || normalizeIso(sheetValue as any) || '';
};

const defaultServiceTypeForDate = (date: Date): string => {
  const nth = Math.floor((date.getDate() - 1) / 7) + 1;
  return (nth === 1 || nth === 3 || nth === 5) ? 'Communion' : 'Offering';
};

const todayISO = () => {
  const tz = (Session.getScriptTimeZone && Session.getScriptTimeZone()) || 'Etc/UTC';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
};

const defaultYouTubeBackfillStartDate = () => {
  const today = new Date();
  return `${today.getFullYear() - 1}-10-01`;
};

function ensureServiceColumns(
  sh: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[],
  required: string[]
) {
  let lastCol = sh.getLastColumn();
  const normalized = new Set(headers.map(h => String(h || '').trim().toLowerCase()));
  for (const name of required) {
    const key = String(name || '').trim().toLowerCase();
    if (normalized.has(key)) continue;
    sh.insertColumnAfter(lastCol);
    lastCol = sh.getLastColumn();
    sh.getRange(1, lastCol).setValue(name);
    headers.push(name);
    normalized.add(key);
  }
  return lastCol;
}

function normalizeSongLookup(value: unknown): string {
  return String(value || '')
    .toLowerCase()
    .replace(/\([^)]*\)|\[[^\]]*\]/g, ' ')
    .replace(/\+sp|\+es/g, ' ')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function looksLikeSongItemType(value: unknown) {
  const text = String(value || '').trim().toLowerCase();
  if (!text) return false;
  return text.includes('song') || text.includes('worship');
}

function formatPerformanceLabel(service: Pick<ServiceItem, 'date' | 'time' | 'type'>) {
  const parts = [String(service.date || '').trim(), String(service.time || '').trim()].filter(Boolean);
  const type = String(service.type || '').trim();
  if (type) parts.push(type);
  return parts.join(' - ');
}

function formatSecondsAsTimestamp(totalSeconds: number) {
  const seconds = Math.max(0, Math.floor(Number(totalSeconds) || 0));
  const hours = Math.floor(seconds / 3600);
  const minutes = Math.floor((seconds % 3600) / 60);
  const remainder = seconds % 60;
  if (hours > 0) return `${hours}:${String(minutes).padStart(2, '0')}:${String(remainder).padStart(2, '0')}`;
  return `${minutes}:${String(remainder).padStart(2, '0')}`;
}

function appendYouTubeStartTime(url: string, startSeconds: number) {
  const base = String(url || '').trim();
  const seconds = Math.max(0, Math.floor(Number(startSeconds) || 0));
  if (!base || !seconds) return base;
  const joiner = base.includes('?') ? '&' : '?';
  if (/[?&]t=\d+s?\b/i.test(base) || /[?&]start=\d+\b/i.test(base)) return base;
  return `${base}${joiner}t=${seconds}s`;
}

type SongPerformanceLinkRow = {
  serviceId: string;
  youtubeUrl: string;
  startSeconds: number;
  startLabel: string;
};

type SaveSongPerformanceTimestampInput = {
  songName?: string;
  serviceId?: string;
  youtubeUrl?: string;
  timestampInput?: string | number;
};

function getSongPerformanceLinkMap(songName: string) {
  const target = normalizeSongLookup(songName);
  const empty = new Map<string, SongPerformanceLinkRow>();
  if (!target) return empty;

  const headers = Object.values(SONG_PERFORMANCES_COL);
  const sh = getOrCreateSheet(SONG_PERFORMANCES_SHEET, headers);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return empty;

  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headerRow.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const songNameIdx = col(SONG_PERFORMANCES_COL.songName);
  const serviceIdIdx = col(SONG_PERFORMANCES_COL.serviceId);
  const youtubeUrlIdx = col(SONG_PERFORMANCES_COL.youtubeUrl);
  const startSecondsIdx = col(SONG_PERFORMANCES_COL.startSeconds);
  const startLabelIdx = col(SONG_PERFORMANCES_COL.startLabel);
  if (songNameIdx < 0 || serviceIdIdx < 0) return empty;

  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const byServiceId = new Map<string, SongPerformanceLinkRow>();
  for (const row of rows) {
    const rowSongName = String(row[songNameIdx] ?? '').trim();
    if (!rowSongName || normalizeSongLookup(rowSongName) !== target) continue;
    const serviceId = String(row[serviceIdIdx] ?? '').trim();
    if (!serviceId) continue;
    const youtubeUrl = youtubeUrlIdx >= 0 ? String(row[youtubeUrlIdx] ?? '').trim() : '';
    const numericStart = startSecondsIdx >= 0 ? Math.max(0, Math.floor(Number(row[startSecondsIdx]) || 0)) : 0;
    byServiceId.set(serviceId, {
      serviceId,
      youtubeUrl,
      startSeconds: numericStart,
      startLabel: numericStart > 0 ? formatSecondsAsTimestamp(numericStart) : ''
    });
  }
  return byServiceId;
}

export function getSongPerformancePlayback(songName: string, serviceId: string, fallbackUrl?: string) {
  const normalizedServiceId = String(serviceId || '').trim();
  if (!songName || !normalizedServiceId) {
    return { youtubeUrl: '', baseYoutubeUrl: '', startSeconds: 0, startLabel: '' };
  }
  const row = getSongPerformanceLinkMap(songName).get(normalizedServiceId);
  const baseYoutubeUrl = String(row?.youtubeUrl || fallbackUrl || '').trim();
  const startSeconds = Math.max(0, Math.floor(Number(row?.startSeconds) || 0));
  return {
    youtubeUrl: appendYouTubeStartTime(baseYoutubeUrl, startSeconds),
    baseYoutubeUrl,
    startSeconds,
    startLabel: startSeconds > 0 ? formatSecondsAsTimestamp(startSeconds) : ''
  };
}

function extractVideoIdFromYouTubeUrl(url: string) {
  const text = String(url || '').trim();
  if (!text) return '';
  const watchMatch = text.match(/[?&]v=([A-Za-z0-9_-]{6,})/);
  if (watchMatch?.[1]) return watchMatch[1];
  const shortMatch = text.match(/youtu\.be\/([A-Za-z0-9_-]{6,})/i);
  if (shortMatch?.[1]) return shortMatch[1];
  const embedMatch = text.match(/\/embed\/([A-Za-z0-9_-]{6,})/i);
  if (embedMatch?.[1]) return embedMatch[1];
  return '';
}

function parseTimestampToSeconds(input: unknown) {
  const raw = String(input ?? '').trim();
  if (!raw) return 0;

  const urlTimeMatch =
    raw.match(/[?&]t=(\d+)s?\b/i) ||
    raw.match(/[?&]start=(\d+)\b/i) ||
    raw.match(/[?&]time_continue=(\d+)\b/i);
  if (urlTimeMatch?.[1]) return Math.max(0, Math.floor(Number(urlTimeMatch[1]) || 0));

  if (/^\d+$/.test(raw)) return Math.max(0, Math.floor(Number(raw) || 0));

  const parts = raw.split(':').map(part => part.trim()).filter(Boolean);
  if (parts.length >= 2 && parts.length <= 3 && parts.every(part => /^\d+$/.test(part))) {
    const nums = parts.map(part => Number(part));
    if (parts.length === 2) return nums[0] * 60 + nums[1];
    return nums[0] * 3600 + nums[1] * 60 + nums[2];
  }

  return 0;
}

function formatIsoDateLabel(iso: string, style: 'long' | 'short' = 'long') {
  const date = dateFromISO(iso);
  if (!date) return iso;
  const month = Utilities.formatDate(date, 'Etc/UTC', style === 'long' ? 'MMMM' : 'MMM');
  const day = String(date.getDate());
  const year = String(date.getFullYear());
  return `${day} ${month} ${year}`;
}

function parseXmlSafe(input: string) {
  try {
    return XmlService.parse(input);
  } catch (_) {
    return null;
  }
}

function getYouTubeChannelId() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(`${YOUTUBE_HELPER_CACHE_KEY}:channel-id`);
  if (cached) return cached;
  const response = UrlFetchApp.fetch(DEFAULT_YOUTUBE_CHANNEL_URL, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Google-Apps-Script)' },
    muteHttpExceptions: true
  });
  const html = response.getContentText();
  const match = html.match(/"channelId":"(UC[^"]+)"/) || html.match(/"externalId":"(UC[^"]+)"/);
  const channelId = match && match[1] ? String(match[1]).trim() : '';
  if (channelId) {
    try { cache.put(`${YOUTUBE_HELPER_CACHE_KEY}:channel-id`, channelId, 21600); } catch (_) {}
  }
  return channelId;
}

function getOrCreateSheet(name: string, headers?: string[]) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers?.length) {
    const lastCol = Math.max(sh.getLastColumn(), headers.length);
    const existing = lastCol > 0
      ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim())
      : [];
    if (!existing.length || !existing.some(Boolean)) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      ensureServiceColumns(sh, existing, headers);
    }
  }
  return sh;
}

function getYouTubeStreamUrlByServiceIdMap() {
  const map = new Map<string, string>();
  let sh: GoogleAppsScript.Spreadsheet.Sheet | null = null;
  try {
    sh = getSheetByName(YOUTUBE_STREAMS_SHEET);
  } catch (_) {
    sh = null;
  }
  if (!sh) return map;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return map;
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const serviceIdIdx = col(YOUTUBE_STREAMS_COL.matchedServiceId);
  const urlIdx = col(YOUTUBE_STREAMS_COL.url);
  if (serviceIdIdx < 0 || urlIdx < 0) return map;
  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  rows.forEach((row) => {
    const serviceId = String(row[serviceIdIdx] ?? '').trim();
    const url = String(row[urlIdx] ?? '').trim();
    if (serviceId && url && !map.has(serviceId)) map.set(serviceId, url);
  });
  return map;
}

function extractYouTubeVideoId(url: string) {
  const text = String(url || '').trim();
  if (!text) return '';
  const match =
    text.match(/[?&]v=([^&#]+)/i) ||
    text.match(/youtu\.be\/([^?&#/]+)/i) ||
    text.match(/\/watch\/([^?&#/]+)/i);
  return match && match[1] ? String(match[1]).trim() : '';
}

function fetchYouTubeFeedEntries() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(`${YOUTUBE_HELPER_CACHE_KEY}:feed`);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed)) return parsed;
    } catch (_) {}
  }
  const channelId = getYouTubeChannelId();
  if (!channelId) return [];
  const response = UrlFetchApp.fetch(`https://www.youtube.com/feeds/videos.xml?channel_id=${encodeURIComponent(channelId)}`, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Google-Apps-Script)' },
    muteHttpExceptions: true
  });
  const xml = response.getContentText();
  const doc = parseXmlSafe(xml);
  if (!doc) return [];
  const root = doc.getRootElement();
  const atomNs = root.getNamespace();
  const entries = root.getChildren('entry', atomNs);
  const items = entries.map((entry: GoogleAppsScript.XML_Service.Element) => {
    const title = String(entry.getChildText('title', atomNs) || '').trim();
    const published = String(entry.getChildText('published', atomNs) || '').trim();
    const links = entry.getChildren('link', atomNs);
    let url = '';
    for (const link of links) {
      const href = String(link.getAttribute('href')?.getValue() || '').trim();
      const rel = String(link.getAttribute('rel')?.getValue() || '').trim();
      if (href && (!rel || rel === 'alternate')) {
        url = href;
        break;
      }
    }
    return { title, published, url };
  }).filter(item => item.title && item.url);
  try { cache.put(`${YOUTUBE_HELPER_CACHE_KEY}:feed`, JSON.stringify(items), 1800); } catch (_) {}
  return items;
}

function decodeJsonStringLiteral(value: string) {
  try {
    return JSON.parse(`"${String(value || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`);
  } catch (_) {
    return String(value || '')
      .replace(/\\u0026/g, '&')
      .replace(/\\u003d/g, '=')
      .replace(/\\u002f/g, '/')
      .replace(/\\"/g, '"')
      .replace(/\\\\/g, '\\');
  }
}

function extractJsonObjectAfterMarker(html: string, marker: string) {
  const source = String(html || '');
  const startIdx = source.indexOf(marker);
  if (startIdx < 0) return '';
  const braceStart = source.indexOf('{', startIdx + marker.length);
  if (braceStart < 0) return '';
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let i = braceStart; i < source.length; i++) {
    const ch = source[i];
    if (inString) {
      if (escaped) {
        escaped = false;
      } else if (ch === '\\') {
        escaped = true;
      } else if (ch === '"') {
        inString = false;
      }
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === '{') depth += 1;
    else if (ch === '}') {
      depth -= 1;
      if (depth === 0) {
        return source.slice(braceStart, i + 1);
      }
    }
  }
  return '';
}

function extractJsonParseStringAfterMarker(html: string, marker: string) {
  const source = String(html || '');
  const startIdx = source.indexOf(marker);
  if (startIdx < 0) return '';
  const parseIdx = source.indexOf('JSON.parse(', startIdx + marker.length);
  if (parseIdx < 0) return '';
  let idx = parseIdx + 'JSON.parse('.length;
  while (idx < source.length && /\s/.test(source[idx])) idx += 1;
  const quote = source[idx];
  if (quote !== "'" && quote !== '"') return '';
  idx += 1;
  let out = '';
  for (let i = idx; i < source.length; i++) {
    const ch = source[i];
    if (ch === quote) return out;
    if (ch !== '\\') {
      out += ch;
      continue;
    }
    i += 1;
    if (i >= source.length) break;
    const esc = source[i];
    switch (esc) {
      case '\\': out += '\\'; break;
      case '/': out += '/'; break;
      case "'": out += "'"; break;
      case '"': out += '"'; break;
      case 'b': out += '\b'; break;
      case 'f': out += '\f'; break;
      case 'n': out += '\n'; break;
      case 'r': out += '\r'; break;
      case 't': out += '\t'; break;
      case 'v': out += '\v'; break;
      case '0': out += '\0'; break;
      case 'x': {
        const hex = source.slice(i + 1, i + 3);
        if (/^[0-9a-fA-F]{2}$/.test(hex)) {
          out += String.fromCharCode(parseInt(hex, 16));
          i += 2;
        } else {
          out += 'x';
        }
        break;
      }
      case 'u': {
        const hex = source.slice(i + 1, i + 5);
        if (/^[0-9a-fA-F]{4}$/.test(hex)) {
          out += String.fromCharCode(parseInt(hex, 16));
          i += 4;
        } else {
          out += 'u';
        }
        break;
      }
      case '\n':
      case '\r':
        break;
      default:
        out += esc;
        break;
    }
  }
  return '';
}

function extractQuotedJsStringAfterMarker(html: string, marker: string) {
  const source = String(html || '');
  const startIdx = source.indexOf(marker);
  if (startIdx < 0) return '';
  let idx = startIdx + marker.length;
  while (idx < source.length && /\s/.test(source[idx])) idx += 1;
  const quote = source[idx];
  if (quote !== "'" && quote !== '"') return '';
  idx += 1;
  let out = '';
  for (let i = idx; i < source.length; i++) {
    const ch = source[i];
    if (ch === quote) return out;
    if (ch !== '\\') {
      out += ch;
      continue;
    }
    i += 1;
    if (i >= source.length) break;
    const esc = source[i];
    switch (esc) {
      case '\\': out += '\\'; break;
      case '/': out += '/'; break;
      case "'": out += "'"; break;
      case '"': out += '"'; break;
      case 'b': out += '\b'; break;
      case 'f': out += '\f'; break;
      case 'n': out += '\n'; break;
      case 'r': out += '\r'; break;
      case 't': out += '\t'; break;
      case 'v': out += '\v'; break;
      case '0': out += '\0'; break;
      case 'x': {
        const hex = source.slice(i + 1, i + 3);
        if (/^[0-9a-fA-F]{2}$/.test(hex)) {
          out += String.fromCharCode(parseInt(hex, 16));
          i += 2;
        } else {
          out += 'x';
        }
        break;
      }
      case 'u': {
        const hex = source.slice(i + 1, i + 5);
        if (/^[0-9a-fA-F]{4}$/.test(hex)) {
          out += String.fromCharCode(parseInt(hex, 16));
          i += 4;
        } else {
          out += 'u';
        }
        break;
      }
      case '\n':
      case '\r':
        break;
      default:
        out += esc;
        break;
    }
  }
  return '';
}

function loadYouTubeJsonBlob(html: string, markers: string[]) {
  const source = String(html || '');
  for (const marker of markers) {
    const parsedString = extractJsonParseStringAfterMarker(source, marker);
    if (parsedString) {
      try { return JSON.parse(parsedString); } catch (_) {}
    }
    const quotedString = extractQuotedJsStringAfterMarker(source, marker);
    if (quotedString) {
      try { return JSON.parse(quotedString); } catch (_) {}
    }
    const rawObject = extractJsonObjectAfterMarker(source, marker);
    if (rawObject) {
      try { return JSON.parse(rawObject); } catch (_) {}
    }
  }
  return null;
}

function loadYouTubeInitialData(html: string) {
  return loadYouTubeJsonBlob(html, [
    'var ytInitialData = ',
    'window["ytInitialData"] = ',
    'ytInitialData = '
  ]);
}

function collectYouTubeEntriesFromInitialData(node: any, sink: Array<{ title: string; url: string; published: string }>, seen: Set<string>) {
  if (!node || typeof node !== 'object') return;
  const maybeRenderer =
    node.videoRenderer ||
    node.gridVideoRenderer ||
    node.richItemRenderer?.content?.videoRenderer ||
    node.playlistVideoRenderer ||
    null;
  if (maybeRenderer && typeof maybeRenderer === 'object') {
    const videoId = String(maybeRenderer.videoId || '').trim();
    const title =
      String(maybeRenderer.title?.runs?.[0]?.text || maybeRenderer.title?.simpleText || '').trim();
    const published =
      String(maybeRenderer.publishedTimeText?.simpleText || maybeRenderer.publishedTimeText?.runs?.[0]?.text || '').trim();
    if (videoId && title && !seen.has(videoId)) {
      seen.add(videoId);
      sink.push({
        title: decodeJsonStringLiteral(title),
        url: `https://www.youtube.com/watch?v=${videoId}`,
        published
      });
    }
  }
  if (Array.isArray(node)) {
    node.forEach(child => collectYouTubeEntriesFromInitialData(child, sink, seen));
    return;
  }
  Object.keys(node).forEach(key => {
    try {
      collectYouTubeEntriesFromInitialData((node as any)[key], sink, seen);
    } catch (_) {}
  });
}

function stripHtmlTags(input?: string) {
  return decodeHtmlEntities(String(input || '').replace(/<[^>]+>/g, ' '))
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeYouTubeTitleCandidate(input?: string) {
  let text = stripHtmlTags(input);
  if (!text) return '';
  text = text.replace(/\s+\d+\s+(?:hour|hours|minute|minutes|second|seconds)\b.*$/i, '').trim();
  text = text.replace(/\s+\d+:\d+(?::\d+)?$/i, '').trim();
  return text;
}

function extractYouTubeEntriesViaAnchors(html: string) {
  const items: Array<{ videoId: string; title: string; published: string }> = [];
  const seen = new Set<string>();
  const source = String(html || '');
  const anchorRegex = /<a\b([^>]*?)href=(['"])\/watch\?v=([^"'&]+)[^'"]*\2([^>]*)>([\s\S]*?)<\/a>/gi;
  let match: RegExpExecArray | null;
  while ((match = anchorRegex.exec(source))) {
    const attrs = `${match[1] || ''} ${match[4] || ''}`;
    const videoId = String(match[3] || '').trim();
    if (!videoId || seen.has(videoId)) continue;
    const ariaLabelMatch = attrs.match(/\baria-label=(['"])([\s\S]*?)\1/i);
    const titleAttrMatch = attrs.match(/\btitle=(['"])([\s\S]*?)\1/i);
    const titleFromAttr = ariaLabelMatch?.[2] || titleAttrMatch?.[2] || '';
    const title = normalizeYouTubeTitleCandidate(titleFromAttr || match[5] || '');
    if (!title) continue;
    seen.add(videoId);
    items.push({
      videoId,
      title,
      published: ''
    });
  }
  return items;
}

function extractYouTubeEntriesViaWatchSnippets(html: string) {
  const items: Array<{ videoId: string; title: string; published: string }> = [];
  const seen = new Set<string>();
  const source = String(html || '');
  const watchRegex = /\/watch\?v=([^"'&\\]+)[^"'\\<]{0,4000}/g;
  let match: RegExpExecArray | null;
  while ((match = watchRegex.exec(source))) {
    const videoId = String(match[1] || '').trim();
    if (!videoId || seen.has(videoId)) continue;
    const windowStart = Math.max(0, match.index - 800);
    const windowEnd = Math.min(source.length, match.index + 5000);
    const snippet = source.slice(windowStart, windowEnd);
    const ariaLabelMatch = snippet.match(/\baria-label=(['"])([\s\S]{1,400}?)\1/i);
    const titleTextMatch =
      snippet.match(/<span\b[^>]*>([^<]{3,200})<\/span>/i) ||
      snippet.match(/"title"\s*:\s*\{"runs":\[\{"text":"([^"]{3,200})"/i) ||
      snippet.match(/"simpleText":"([^"]{3,200})"/i);
    const publishedMatch =
      snippet.match(/"publishedTimeText"\s*:\s*\{"simpleText":"([^"]{1,120})"/i) ||
      snippet.match(/Streamed live on[^<]{0,80}/i);
    const title = normalizeYouTubeTitleCandidate((ariaLabelMatch?.[2] || titleTextMatch?.[1] || ''));
    if (!title) continue;
    seen.add(videoId);
    items.push({
      videoId,
      title,
      published: stripHtmlTags(publishedMatch?.[0] || publishedMatch?.[1] || '')
    });
  }
  return items;
}

function extractYouTubeEntriesFromJsonText(jsonText: string) {
  const items: Array<{ videoId: string; title: string; published: string }> = [];
  const seen = new Set<string>();
  const source = String(jsonText || '');
  if (!source) return items;

  const add = (videoId: string, rawTitle: string, rawPublished = '') => {
    const id = String(videoId || '').trim();
    const title = normalizeYouTubeTitleCandidate(rawTitle);
    if (!id || !title || seen.has(id)) return;
    seen.add(id);
    items.push({
      videoId: id,
      title,
      published: stripHtmlTags(rawPublished)
    });
  };

  const videoIdRegex = /"videoId":"([^"]+)"/g;
  let match: RegExpExecArray | null;
  while ((match = videoIdRegex.exec(source))) {
    const videoId = String(match[1] || '').trim();
    if (!videoId || seen.has(videoId)) continue;
    const windowStart = Math.max(0, match.index - 500);
    const windowEnd = Math.min(source.length, match.index + 6000);
    const snippet = source.slice(windowStart, windowEnd);

    const titleMatch =
      snippet.match(/"(?:title|headline)":\{"runs":\[\{"text":"([^"]{3,200})"/) ||
      snippet.match(/"(?:title|headline)":\{"simpleText":"([^"]{3,200})"/) ||
      snippet.match(/"accessibilityData":\{"label":"([^"]{3,300})"/);
    const publishedMatch =
      snippet.match(/"publishedTimeText":\{"simpleText":"([^"]{1,120})"/) ||
      snippet.match(/"publishedTimeText":\{"runs":\[\{"text":"([^"]{1,120})"/) ||
      snippet.match(/"dateText":\{"simpleText":"([^"]{1,120})"/) ||
      snippet.match(/"dateText":\{"runs":\[\{"text":"([^"]{1,120})"/);

    add(
      videoId,
      String(titleMatch?.[1] || ''),
      String(publishedMatch?.[1] || '')
    );
  }

  return items;
}

function inspectYouTubeInitialData(html: string) {
  const source = String(html || '');
  const parseMarkers = [
    'var ytInitialData = ',
    'window["ytInitialData"] = ',
    'ytInitialData = '
  ];
  let parsedString = '';
  let parsedMarker = '';
  parseMarkers.some((marker) => {
    parsedString = extractJsonParseStringAfterMarker(source, marker);
    parsedMarker = parsedString ? marker : '';
    return !!parsedString;
  });

  let quotedString = '';
  let quotedMarker = '';
  if (!parsedString) {
    parseMarkers.some((marker) => {
      quotedString = extractQuotedJsStringAfterMarker(source, marker);
      quotedMarker = quotedString ? marker : '';
      return !!quotedString;
    });
  }

  let parseError = '';
  let parsedValue: any = null;
  if (parsedString) {
    try {
      parsedValue = JSON.parse(parsedString);
    } catch (err) {
      parseError = String((err as any)?.message || err || 'JSON.parse failed');
    }
  }
  if (!parsedValue && quotedString) {
    try {
      parsedValue = JSON.parse(quotedString);
    } catch (err) {
      parseError = parseError || String((err as any)?.message || err || 'Quoted string parse failed');
    }
  }

  let rawObject = '';
  let rawMarker = '';
  if (!parsedValue) {
    parseMarkers.some((marker) => {
      rawObject = extractJsonObjectAfterMarker(source, marker);
      rawMarker = rawObject ? marker : '';
      return !!rawObject;
    });
    if (rawObject) {
      try {
        parsedValue = JSON.parse(rawObject);
      } catch (err) {
        parseError = parseError || String((err as any)?.message || err || 'Raw object parse failed');
      }
    }
  }

  return {
    value: parsedValue,
    parseMarker: parsedMarker || quotedMarker,
    rawMarker,
    parsedStringLength: parsedString ? parsedString.length : 0,
    quotedStringLength: quotedString ? quotedString.length : 0,
    decodedText: parsedString || quotedString || rawObject || '',
    rawObjectLength: rawObject ? rawObject.length : 0,
    parseError,
    snippet: source.slice(
      Math.max(0, source.indexOf(parsedMarker || quotedMarker || rawMarker || 'ytInitialData') - 120),
      Math.min(source.length, source.indexOf(parsedMarker || quotedMarker || rawMarker || 'ytInitialData') + 280)
    ).replace(/\s+/g, ' ')
  };
}

function fetchYouTubeStreamsPageEntries() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(`${YOUTUBE_HELPER_CACHE_KEY}:streams-page`);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed)) return parsed;
    } catch (_) {}
  }

  const response = UrlFetchApp.fetch(DEFAULT_YOUTUBE_STREAMS_URL, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Google-Apps-Script)' },
    muteHttpExceptions: true
  });
  const html = response.getContentText();
  const items = extractYouTubeEntriesFromHtml(html);
  try { cache.put(`${YOUTUBE_HELPER_CACHE_KEY}:streams-page`, JSON.stringify(items), 1800); } catch (_) {}
  return items;
}

function extractYouTubeEntriesFromHtml(html: string) {
  const items: Array<{ title: string; url: string; published: string }> = [];
  const seen = new Set<string>();
  const addItem = (videoId: string, rawTitle: string, published = '') => {
    const id = String(videoId || '').trim();
    const title = decodeJsonStringLiteral(
      String(rawTitle || '')
        .replace(/<[^>]+>/g, ' ')
        .replace(/\s+/g, ' ')
        .trim()
    );
    if (!id || !title || seen.has(id)) return;
    seen.add(id);
    items.push({
      title,
      url: `https://www.youtube.com/watch?v=${id}`,
      published: String(published || '').trim()
    });
  };
  extractYouTubeEntriesViaAnchors(html).forEach(item => addItem(item.videoId, item.title, item.published));
  extractYouTubeEntriesViaWatchSnippets(html).forEach(item => addItem(item.videoId, item.title, item.published));

  let match: RegExpExecArray | null;
  const jsonRegex = /"videoId":"([^"]+)"[\s\S]{0,12000}?"title":\{"runs":\[\{"text":"([^"]+)"\}\]/g;
  while ((match = jsonRegex.exec(html))) {
    addItem(String(match[1] || '').trim(), String(match[2] || '').trim(), '');
  }

  const broadJsonRegex = /"videoId":"([^"]+)"[\s\S]{0,12000}?"(?:title|headline)":\{"(?:runs":\[\{"text":"([^"]+)"\}\]|"simpleText":"([^"]+)")/g;
  while ((match = broadJsonRegex.exec(html))) {
    addItem(String(match[1] || '').trim(), String(match[2] || match[3] || '').trim(), '');
  }

  const publishedJsonRegex = /"videoId":"([^"]+)"[\s\S]{0,12000}?"publishedTimeText":\{"simpleText":"([^"]+)"\}[\s\S]{0,12000}?"(?:title|headline)":\{"(?:runs":\[\{"text":"([^"]+)"\}\]|"simpleText":"([^"]+)")/g;
  while ((match = publishedJsonRegex.exec(html))) {
    addItem(String(match[1] || '').trim(), String(match[3] || match[4] || '').trim(), String(match[2] || '').trim());
  }

  const initialInspection = inspectYouTubeInitialData(html);
  const initialData = initialInspection.value;
  extractYouTubeEntriesFromJsonText(initialInspection.decodedText).forEach(item => addItem(item.videoId, item.title, item.published));
  if (initialData) {
    try {
      collectYouTubeEntriesFromInitialData(initialData, items, seen);
    } catch (_) {
      // Ignore parse failures and keep regex-derived items.
    }
  }
  return items;
}

function fetchYouTubeSearchEntries(query: string) {
  const trimmed = String(query || '').trim();
  if (!trimmed) return [];
  const cache = CacheService.getScriptCache();
  const cacheKey = `${YOUTUBE_HELPER_CACHE_KEY}:search:${Utilities.base64EncodeWebSafe(trimmed).slice(0, 80)}`;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed)) return parsed;
    } catch (_) {}
  }
  const response = UrlFetchApp.fetch(`https://www.youtube.com/results?search_query=${encodeURIComponent(trimmed)}`, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Google-Apps-Script)' },
    muteHttpExceptions: true
  });
  const html = response.getContentText();
  const items = extractYouTubeEntriesFromHtml(html);
  try { cache.put(cacheKey, JSON.stringify(items), 1800); } catch (_) {}
  return items;
}

function summarizeYouTubeHtml(html: string) {
  const text = String(html || '');
  return {
    length: text.length,
    hasWatchHref: /\/watch\?v=/.test(text) || /\\\/watch\\\?v=/.test(text),
    hasConsent: /consent/i.test(text),
    hasChannelId: /"channelId":"UC/.test(text),
    hasDatePublished: /datePublished/.test(text),
    hasStreamedLive: /Streamed live on/i.test(text),
    titleSample: ((text.match(/<title>([^<]+)<\/title>/i) || [])[1] || '').trim(),
    snippet: text.replace(/\s+/g, ' ').slice(0, 400)
  };
}

function fetchYouTubeUrlDebug(url: string) {
  const target = String(url || '').trim();
  if (!target) {
    return {
      url: target,
      status: 0,
      finalUrl: '',
      ...summarizeYouTubeHtml('')
    };
  }
  try {
    const response = UrlFetchApp.fetch(target, {
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Google-Apps-Script)' },
      muteHttpExceptions: true,
      followRedirects: true
    });
    const html = response.getContentText();
    return {
      url: target,
      status: response.getResponseCode(),
      finalUrl: '',
      ...summarizeYouTubeHtml(html)
    };
  } catch (err) {
    return {
      url: target,
      status: -1,
      finalUrl: '',
      ...summarizeYouTubeHtml(''),
      snippet: String((err && (err as any).message) || err || 'Unknown fetch error')
    };
  }
}

function buildYouTubeSearchQueries(isoDate: string) {
  const longLabel = formatIsoDateLabel(isoDate, 'long');
  const shortLabel = formatIsoDateLabel(isoDate, 'short');
  return [
    `${DEFAULT_YOUTUBE_TITLE_PREFIX} ${longLabel}`.trim(),
    `${DEFAULT_YOUTUBE_TITLE_PREFIX} ${shortLabel}`.trim(),
    longLabel,
    shortLabel
  ].filter(Boolean);
}

function listServiceDatesForYouTubeBackfill() {
  const minDate = defaultYouTubeBackfillStartDate();
  const cutoff = todayISO();
  const sh = getSheetByName(SERVICES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [] as string[];
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const idIdx = col(SERVICES_COL.id);
  const dateIdx = col(SERVICES_COL.date);
  if (idIdx < 0 && dateIdx < 0) return [] as string[];
  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const dates = new Set<string>();
  rows.forEach((row) => {
    const serviceId = idIdx >= 0 ? String(row[idIdx] ?? '').trim() : '';
    const date = canonicalServiceDate(serviceId, dateIdx >= 0 ? row[dateIdx] : '');
    if (date && date >= minDate && date <= cutoff) dates.add(date);
  });
  return Array.from(dates).sort();
}

function getYouTubeCatalogCursor() {
  try {
    return String(PropertiesService.getDocumentProperties().getProperty(YOUTUBE_STREAMS_CURSOR_PROPERTY) || '').trim();
  } catch (_) {
    return '';
  }
}

function setYouTubeCatalogCursor(value: string) {
  try {
    const next = String(value || '').trim();
    if (!next) {
      PropertiesService.getDocumentProperties().deleteProperty(YOUTUBE_STREAMS_CURSOR_PROPERTY);
      return;
    }
    PropertiesService.getDocumentProperties().setProperty(YOUTUBE_STREAMS_CURSOR_PROPERTY, next);
  } catch (_) {}
}

function clearYouTubeCatalogCursor() {
  try {
    PropertiesService.getDocumentProperties().deleteProperty(YOUTUBE_STREAMS_CURSOR_PROPERTY);
  } catch (_) {}
}

function getYouTubeApiKey() {
  try {
    return String(PropertiesService.getScriptProperties().getProperty(YOUTUBE_API_KEY_PROPERTY) || '').trim();
  } catch (_) {
    return '';
  }
}

function youtubeApiGet(path: string, params: Record<string, string>) {
  const apiKey = getYouTubeApiKey();
  if (!apiKey) throw new Error(`Script property "${YOUTUBE_API_KEY_PROPERTY}" is missing.`);
  const pairs = Object.entries({ ...params, key: apiKey })
    .filter(([, value]) => String(value || '').trim() !== '')
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(String(value))}`);
  const url = `https://www.googleapis.com/youtube/v3/${path}?${pairs.join('&')}`;
  const response = UrlFetchApp.fetch(url, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Google-Apps-Script)' },
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  const body = response.getContentText();
  if (status < 200 || status >= 300) {
    let message = `YouTube API request failed (${status})`;
    try {
      const parsed = JSON.parse(body);
      const detail = String(parsed?.error?.message || '').trim();
      if (detail) message += `: ${detail}`;
    } catch (_) {}
    throw new Error(message);
  }
  try {
    return JSON.parse(body);
  } catch (err) {
    throw new Error(`Unable to parse YouTube API response for ${path}: ${err}`);
  }
}

function getYouTubeChannelApiContext() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(`${YOUTUBE_HELPER_CACHE_KEY}:channel-api-context`);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (_) {}
  }
  const data = youtubeApiGet('channels', {
    part: 'id,snippet,contentDetails',
    forHandle: YOUTUBE_CHANNEL_HANDLE,
    maxResults: '1'
  });
  const item = Array.isArray(data?.items) ? data.items[0] : null;
  const context = {
    channelId: String(item?.id || '').trim(),
    channelName: String(item?.snippet?.title || '').trim(),
    uploadsPlaylistId: String(item?.contentDetails?.relatedPlaylists?.uploads || '').trim()
  };
  if (!context.channelId || !context.uploadsPlaylistId) {
    throw new Error(`Unable to resolve YouTube channel context for ${YOUTUBE_CHANNEL_HANDLE}.`);
  }
  try { cache.put(`${YOUTUBE_HELPER_CACHE_KEY}:channel-api-context`, JSON.stringify(context), 21600); } catch (_) {}
  return context;
}

function fetchYouTubeUploadsPage(pageToken?: string) {
  const context = getYouTubeChannelApiContext();
  return youtubeApiGet('playlistItems', {
    part: 'snippet,contentDetails,status',
    playlistId: context.uploadsPlaylistId,
    maxResults: '50',
    pageToken: String(pageToken || '').trim()
  });
}

function fetchYouTubeVideosByIds(videoIds: string[]) {
  const ids = Array.from(new Set((videoIds || []).map(id => String(id || '').trim()).filter(Boolean)));
  if (!ids.length) return [] as any[];
  const out: any[] = [];
  for (let i = 0; i < ids.length; i += 50) {
    const chunk = ids.slice(i, i + 50);
    const data = youtubeApiGet('videos', {
      part: 'snippet,liveStreamingDetails,status,contentDetails',
      id: chunk.join(','),
      maxResults: String(chunk.length)
    });
    if (Array.isArray(data?.items)) out.push(...data.items);
  }
  return out;
}

function buildCompactIsoDateLabels(isoDate: string) {
  const date = dateFromISO(isoDate);
  if (!date) return [] as string[];
  const day = String(date.getDate());
  const monthShort = Utilities.formatDate(date, 'Etc/UTC', 'MMM');
  const year2 = String(date.getFullYear()).slice(-2);
  return [
    `${day}${monthShort}${year2}`,
    `${day} ${monthShort} ${year2}`,
    `${day}${monthShort}${String(date.getFullYear())}`,
    `${day} ${monthShort} ${String(date.getFullYear())}`
  ].map(v => v.toLowerCase());
}

function fetchYouTubeCandidateEntries(isoDate?: string) {
  const byUrl = new Map<string, { title: string; url: string; published: string; sourceRank: number }>();
  const merge = (entry: { title: string; url: string; published: string }, sourceRank: number) => {
    const url = String(entry?.url || '').trim();
    const title = String(entry?.title || '').trim();
    if (!url || !title) return;
    const existing = byUrl.get(url);
    if (!existing) {
      byUrl.set(url, { title, url, published: String(entry?.published || '').trim(), sourceRank });
      return;
    }
    if (!existing.published && entry.published) existing.published = String(entry.published).trim();
    if (sourceRank > existing.sourceRank && title) {
      existing.title = title;
      existing.sourceRank = sourceRank;
    } else if (sourceRank === existing.sourceRank && (!existing.title || existing.title.length < title.length) && title) {
      existing.title = title;
    }
  };
  fetchYouTubeFeedEntries().forEach(item => merge(item, 3));
  fetchYouTubeStreamsPageEntries().forEach(item => merge(item, 4));
  return Array.from(byUrl.values()).map(({ sourceRank, ...item }) => item);
}

function fetchYouTubeVideoMetadata(url: string) {
  const target = String(url || '').trim();
  if (!target) return { published: '', title: '', channelId: '', channelName: '' };
  const cache = CacheService.getScriptCache();
  const cacheKey = `${YOUTUBE_HELPER_CACHE_KEY}:video:${Utilities.base64EncodeWebSafe(target).slice(0, 80)}`;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      return {
        published: String(parsed?.published || '').trim(),
        title: String(parsed?.title || '').trim(),
        channelId: String(parsed?.channelId || '').trim(),
        channelName: String(parsed?.channelName || '').trim()
      };
    } catch (_) {}
  }
  const response = UrlFetchApp.fetch(target, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Google-Apps-Script)' },
    muteHttpExceptions: true
  });
  const html = response.getContentText();
  const playerResponse = loadYouTubeJsonBlob(html, [
    'var ytInitialPlayerResponse = ',
    'window["ytInitialPlayerResponse"] = ',
    'ytInitialPlayerResponse = '
  ]);
  const parseNamedMonthDate = (raw: string) => {
    const text = String(raw || '').trim();
    if (!text) return '';
    const parsed = new Date(text);
    if (isNaN(parsed.getTime())) return '';
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, '0');
    const d = String(parsed.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  };
  const streamedLiveRaw =
    String(playerResponse?.microformat?.playerMicroformatRenderer?.liveBroadcastDetails?.startTimestamp || '').trim() ||
    ((html.match(/Streamed live on ([A-Za-z]{3,9}\s+\d{1,2},\s+\d{4})/) || [])[1]) ||
    ((html.match(/"dateText":\{"simpleText":"Streamed live on ([^"]+)"\}/) || [])[1]) ||
    '';
  const published =
    normalizeIso(String(playerResponse?.microformat?.playerMicroformatRenderer?.publishDate || '').trim() as any) ||
    normalizeIso(String(playerResponse?.microformat?.playerMicroformatRenderer?.uploadDate || '').trim() as any) ||
    normalizeIso(String(streamedLiveRaw).slice(0, 10) as any) ||
    parseNamedMonthDate(streamedLiveRaw) ||
    ((html.match(/"datePublished":"(\d{4}-\d{2}-\d{2})"/) || [])[1]) ||
    ((html.match(/<meta\s+itemprop="datePublished"\s+content="(\d{4}-\d{2}-\d{2})"/i) || [])[1]) ||
    ((html.match(/<meta\s+property="og:video:release_date"\s+content="(\d{4}-\d{2}-\d{2})/i) || [])[1]) ||
    ((html.match(/itemprop="datePublished"\s+content="(\d{4}-\d{2}-\d{2})"/) || [])[1]) ||
    '';
  const title =
    decodeJsonStringLiteral(String(playerResponse?.videoDetails?.title || '').trim()) ||
    decodeJsonStringLiteral((((html.match(/<meta\s+property="og:title"\s+content="([^"]+)"/i) || [])[1]) || '').trim()) ||
    decodeJsonStringLiteral((((html.match(/<meta\s+name="title"\s+content="([^"]+)"/i) || [])[1]) || '').trim()) ||
    decodeJsonStringLiteral((((html.match(/"title":"([^"]+)"/) || [])[1]) || '').trim());
  const channelId =
    String(playerResponse?.videoDetails?.channelId || '').trim() ||
    String(playerResponse?.microformat?.playerMicroformatRenderer?.externalChannelId || '').trim() ||
    ((html.match(/<meta\s+itemprop="channelId"\s+content="(UC[^"]+)"/i) || [])[1]) ||
    ((html.match(/"channelId":"(UC[^"]+)"/) || [])[1]) ||
    ((html.match(/"externalChannelId":"(UC[^"]+)"/) || [])[1]) ||
    '';
  const channelName =
    decodeJsonStringLiteral(String(playerResponse?.videoDetails?.author || '').trim()) ||
    decodeJsonStringLiteral(String(playerResponse?.microformat?.playerMicroformatRenderer?.ownerChannelName || '').trim()) ||
    decodeJsonStringLiteral((((html.match(/<meta\s+itemprop="author"\s+content="([^"]+)"/i) || [])[1]) || '').trim()) ||
    decodeJsonStringLiteral((((html.match(/<link\s+itemprop="name"\s+content="([^"]+)"/i) || [])[1]) || '').trim()) ||
    decodeJsonStringLiteral((((html.match(/"ownerChannelName":"([^"]+)"/) || [])[1]) || '').trim()) ||
    decodeJsonStringLiteral((((html.match(/"channelName":"([^"]+)"/) || [])[1]) || '').trim());
  const result = {
    published: String(published || '').trim(),
    title: String(title || '').trim(),
    channelId: String(channelId || '').trim(),
    channelName: String(channelName || '').trim()
  };
  try { cache.put(cacheKey, JSON.stringify(result), 21600); } catch (_) {}
  return result;
}

function inferIsoDateFromYouTubeTitle(title: string) {
  const text = String(title || '').trim();
  if (!text) return '';
  const explicit =
    text.match(/\b(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})\b/) ||
    text.match(/\b(\d{1,2})([A-Za-z]{3,9})(\d{2,4})\b/);
  if (explicit) {
    const day = Number(explicit[1]);
    const monthText = String(explicit[2] || '').slice(0, 3);
    const yearRaw = String(explicit[3] || '');
    const monthMap: Record<string, number> = {
      jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
      jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
    };
    const month = monthMap[monthText.toLowerCase()];
    let year = Number(yearRaw);
    if (yearRaw.length === 2) year += year >= 70 ? 1900 : 2000;
    if (month && day >= 1 && day <= 31 && year >= 2000) {
      return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }
  return '';
}

function listYouTubeStreamCatalogEntries(options?: { pageToken?: string }) {
  const minDate = defaultYouTubeBackfillStartDate();
  const context = getYouTubeChannelApiContext();
  const page = fetchYouTubeUploadsPage(String(options?.pageToken || '').trim());
  const playlistItems = Array.isArray(page?.items) ? page.items : [];
  const videoIds = playlistItems
    .map((item: any) => String(item?.contentDetails?.videoId || item?.snippet?.resourceId?.videoId || '').trim())
    .filter(Boolean);
  const detailsById = new Map<string, any>();
  fetchYouTubeVideosByIds(videoIds).forEach((item: any) => {
    const id = String(item?.id || '').trim();
    if (id) detailsById.set(id, item);
  });

  const items = playlistItems.map((item: any) => {
    const videoId = String(item?.contentDetails?.videoId || item?.snippet?.resourceId?.videoId || '').trim();
    const detail = detailsById.get(videoId) || {};
    const title = String(detail?.snippet?.title || item?.snippet?.title || '').trim();
    const channelId = String(detail?.snippet?.channelId || context.channelId || '').trim();
    const channelName = String(detail?.snippet?.channelTitle || context.channelName || '').trim();
    const publishedIso =
      normalizeIso(String(detail?.snippet?.publishedAt || '').slice(0, 10) as any) ||
      normalizeIso(String(item?.contentDetails?.videoPublishedAt || '').slice(0, 10) as any) ||
      '';
    const actualStartIso =
      normalizeIso(String(detail?.liveStreamingDetails?.actualStartTime || '').slice(0, 10) as any) ||
      normalizeIso(String(detail?.liveStreamingDetails?.scheduledStartTime || '').slice(0, 10) as any) ||
      '';
    const inferredTitleDate = inferIsoDateFromYouTubeTitle(title);
    const streamDate = actualStartIso || publishedIso || inferredTitleDate || '';
    return {
      videoId,
      url: videoId ? `https://www.youtube.com/watch?v=${videoId}` : '',
      title,
      published: publishedIso,
      channelId,
      channelName,
      source: 'youtube-api',
      streamDate
    };
  }).filter((entry) => {
    if (!entry.videoId || !entry.url || !entry.title) return false;
    if (entry.channelId && entry.channelId !== context.channelId) return false;
    if (entry.streamDate && entry.streamDate < minDate) return false;
    return true;
  });

  const oldestStreamDate = items.length ? items[items.length - 1].streamDate : '';
  const nextPageToken = String(page?.nextPageToken || '').trim();
  const reachedEnd = !nextPageToken || (!!oldestStreamDate && oldestStreamDate < minDate);

  return {
    items,
    processedEntries: playlistItems.length,
    totalEntries: Number(page?.pageInfo?.totalResults || 0) || 0,
    nextPageToken: reachedEnd ? '' : nextPageToken,
    reachedEnd
  };
}

function collectScoredYouTubeCandidates(isoDate: string, enrichMetadata = false) {
  const entries = fetchYouTubeCandidateEntries(isoDate);
  let scored = entries
    .map(item => scoreYouTubeCandidate(item, isoDate))
    .sort((a, b) => b.score - a.score || a.title.localeCompare(b.title));
  if (!enrichMetadata) return scored.filter(item => item.score > 0);
  if (!scored.length) return scored;

  const enriched: YouTubeCandidate[] = [];
  const limit = Math.min(12, scored.length);
  for (let i = 0; i < limit; i++) {
    const base = scored[i];
    const meta = fetchYouTubeVideoMetadata(base.url);
    const rescored = scoreYouTubeCandidate({
      title: meta.title || base.title,
      url: base.url,
      published: meta.published || base.published,
      channelId: meta.channelId || base.channelId,
      channelName: meta.channelName || base.channelName
    }, isoDate);
    enriched.push(rescored);
  }
  for (let i = limit; i < scored.length; i++) enriched.push(scored[i]);
  enriched.sort((a, b) => b.score - a.score || a.title.localeCompare(b.title));
  return enriched.filter(item => item.score > 0);
}

function scoreYouTubeCandidate(item: { title: string; url: string; published: string; channelId?: string; channelName?: string }, isoDate: string): YouTubeCandidate {
  const title = String(item.title || '').trim();
  const published = String(item.published || '').trim();
  const url = String(item.url || '').trim();
  const channelId = String(item.channelId || '').trim();
  const channelName = String(item.channelName || '').trim();
  const longLabel = formatIsoDateLabel(isoDate, 'long');
  const shortLabel = formatIsoDateLabel(isoDate, 'short');
  const compactLabels = buildCompactIsoDateLabels(isoDate);
  const publishedIso = published ? String(published).slice(0, 10) : '';
  const titleLower = title.toLowerCase();
  const targetChannelId = getYouTubeChannelId();
  if (channelId && targetChannelId && channelId !== targetChannelId) {
    return { title, url, published, channelId, channelName, score: 0, reason: `Wrong channel: ${channelName || channelId}` };
  }
  let score = 0;
  let reason = '';
  if (publishedIso === isoDate) {
    score = 120;
    reason = `Video page streamed date matches ${isoDate}`;
  } else if (longLabel && titleLower.includes(longLabel.toLowerCase())) {
    score = 100;
    reason = `Title matches ${longLabel}`;
  } else if (shortLabel && titleLower.includes(shortLabel.toLowerCase())) {
    score = 92;
    reason = `Title matches ${shortLabel}`;
  } else if (compactLabels.some(label => titleLower.includes(label))) {
    score = 88;
    reason = `Title matches compact date for ${isoDate}`;
  } else if (publishedIso) {
    const targetDate = dateFromISO(isoDate);
    const publishedDate = dateFromISO(publishedIso);
    if (targetDate && publishedDate) {
      const diffDays = Math.round((publishedDate.getTime() - targetDate.getTime()) / (24 * 60 * 60 * 1000));
      if (Math.abs(diffDays) <= 3) {
        score = 50 - Math.abs(diffDays) * 10;
        reason = `Published ${Math.abs(diffDays)} day${Math.abs(diffDays) === 1 ? '' : 's'} from ${isoDate}`;
      }
    }
  }
  return { title, url, published, channelId, channelName, score, reason };
}

export type SongPerformance = {
  serviceId: string;
  date: string;
  time: string;
  type: string;
  label: string;
  youtubeUrl: string;
  baseYoutubeUrl?: string;
  startSeconds?: number;
  startLabel?: string;
};

type YouTubeCandidate = {
  title: string;
  url: string;
  published: string;
  channelId?: string;
  channelName?: string;
  score: number;
  reason: string;
};

type SuggestYouTubeStreamResult = {
  url: string;
  title: string;
  matchType: 'exact-title' | 'same-day' | 'nearby' | 'none';
  message: string;
  candidates: Array<{ title: string; url: string; published: string; reason: string }>;
};

export function addService(input: AddServiceInput) {
  const sh = getSheetByName(SERVICES_SHEET);
  const spreadsheetTz = (() => {
    try {
      return SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    } catch (_) {
      return Session.getScriptTimeZone?.() || 'Etc/UTC';
    }
  })();

  let lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  lastCol = ensureServiceColumns(sh, headers, [SERVICES_COL.youtubeUrl, SERVICES_COL.suggestedSongs]);
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const idIdx = col(SERVICES_COL.id);
  const dateIdx = col(SERVICES_COL.date);
  const timeIdx = col(SERVICES_COL.time);
  const typeIdx = col(SERVICES_COL.type);
  const youtubeUrlIdx = col(SERVICES_COL.youtubeUrl);
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
  const suggestedSongsIdx = col(SERVICES_COL.suggestedSongs);

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
  if (youtubeUrlIdx >= 0) vals[youtubeUrlIdx] = input.youtubeUrl ?? '';
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
  if (suggestedSongsIdx >= 0) vals[suggestedSongsIdx] = input.suggestedSongs ?? '';

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
  let lastCol = sh.getLastColumn();
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
      const youtubeUrlIdx = col(SERVICES_COL.youtubeUrl);
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
      const suggestedSongsIdx = col(SERVICES_COL.suggestedSongs);

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
            return Utilities.formatDate(v, spreadsheetTz as string, 'h:mm a');
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

      const rows = body.map(r => {
        const rawId = idIdx >= 0 ? String(r[idIdx] ?? '') : '';
        const rawDate = dateIdx >= 0 ? toISO(r[dateIdx]) : '';
        const rawTime = timeIdx >= 0 ? toTime(r[timeIdx]) : '';
        const derivedTime = deriveTimeFromServiceId(rawId);
        return {
          id: rawId,
          date: rawDate || deriveDateFromServiceId(rawId),
          time: derivedTime || rawTime,
          type: typeIdx >= 0 ? String(r[typeIdx] ?? '') : '',
          youtubeUrl: youtubeUrlIdx >= 0 ? String(r[youtubeUrlIdx] ?? '') : '',
          leader: leaderIdx >= 0 ? String(r[leaderIdx] ?? '') : '',
          preacher: preacherIdx >= 0 ? String(r[preacherIdx] ?? '') : '',
          scripture: scriptureIdx >= 0 ? String(r[scriptureIdx] ?? '') : '',
          scriptureText: scriptureTextIdx >= 0 ? String(r[scriptureTextIdx] ?? '') : '',
          theme: themeIdx >= 0 ? String(r[themeIdx] ?? '') : '',
          keywords: keywordsIdx >= 0 ? String(r[keywordsIdx] ?? '') : '',
          notes: notesIdx >= 0 ? String(r[notesIdx] ?? '') : '',
          suggestedSongs: suggestedSongsIdx >= 0 ? String(r[suggestedSongsIdx] ?? '') : ''
        };
      });

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
  const youtubeUrlIdx = col(SERVICES_COL.youtubeUrl);
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
  const suggestedSongsIdx = col(SERVICES_COL.suggestedSongs);

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
        return Utilities.formatDate(v, spreadsheetTz as string, 'h:mm a');
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

  const items = body.map(r => {
    const rawId = idIdx >= 0 ? String(r[idIdx] ?? '') : '';
    const rawDate = dateIdx >= 0 ? toISO(r[dateIdx]) : '';
    const rawTime = timeIdx >= 0 ? toTime(r[timeIdx]) : '';
    const derivedTime = deriveTimeFromServiceId(rawId);
    return {
      id: rawId,
      date: rawDate || deriveDateFromServiceId(rawId),
      time: derivedTime || rawTime,
      type: typeIdx >= 0 ? String(r[typeIdx] ?? '') : '',
      youtubeUrl: youtubeUrlIdx >= 0 ? String(r[youtubeUrlIdx] ?? '') : '',
      leader: leaderIdx >= 0 ? String(r[leaderIdx] ?? '') : '',
      preacher: preacherIdx >= 0 ? String(r[preacherIdx] ?? '') : '',
      scripture: scriptureIdx >= 0 ? String(r[scriptureIdx] ?? '') : '',
      scriptureText: scriptureTextIdx >= 0 ? String(r[scriptureTextIdx] ?? '') : '',
      theme: themeIdx >= 0 ? String(r[themeIdx] ?? '') : '',
      keywords: keywordsIdx >= 0 ? String(r[keywordsIdx] ?? '') : '',
      notes: notesIdx >= 0 ? String(r[notesIdx] ?? '') : '',
      suggestedSongs: suggestedSongsIdx >= 0 ? String(r[suggestedSongsIdx] ?? '') : ''
    };
  });

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
  ensureUpcomingServicesCoverage();
  const all = fetchServicesUnfiltered();
  return { items: applyServiceFilters(all, opts) };
}

function ensureUpcomingServicesCoverage(weeksAhead = AUTO_SERVICE_WEEKS_AHEAD) {
  const desiredWeeks = Math.min(52, Math.max(1, Number(weeksAhead) || AUTO_SERVICE_WEEKS_AHEAD));
  const sh = getSheetByName(SERVICES_SHEET);
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return { created: [] as { id: string; date: string; time: string; type: string }[] };

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const idIdx = col(SERVICES_COL.id);
  if (idIdx === -1) return { created: [] as { id: string; date: string; time: string; type: string }[] };

  const dateIdx = col(SERVICES_COL.date);
  const timeIdx = col(SERVICES_COL.time);
  const typeIdx = col(SERVICES_COL.type);
  const leaderIdx = col(SERVICES_COL.leader);
  const preacherIdx = col(SERVICES_COL.preacher);
  const suggestedSongsIdx = col(SERVICES_COL.suggestedSongs);

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const lastRow = sh.getLastRow();
    const existing = new Set<string>();
    let latestExisting: Date | null = null;

    if (lastRow >= 2) {
      const idValues = sh.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
      const dateValues = dateIdx >= 0 ? sh.getRange(2, dateIdx + 1, lastRow - 1, 1).getValues() : [];
      for (let i = 0; i < idValues.length; i++) {
        const serviceId = String(idValues[i]?.[0] ?? '').trim();
        if (!serviceId) continue;
        existing.add(serviceId);
        const iso = canonicalServiceDate(serviceId, dateIdx >= 0 ? dateValues[i]?.[0] : '');
        const parsed = dateFromISO(iso);
        if (parsed && (!latestExisting || parsed > latestExisting)) latestExisting = parsed;
      }
    }

    const today = new Date();
    const firstTargetSunday = nextSundayOnOrAfter(today);
    const horizonDate = addDays(firstTargetSunday, (desiredWeeks - 1) * 7);
    const startDate = latestExisting ? addDays(latestExisting, 7) : firstTargetSunday;
    if (startDate > horizonDate) return { created: [] as { id: string; date: string; time: string; type: string }[] };

    const rows: any[][] = [];
    const created: { id: string; date: string; time: string; type: string }[] = [];
    for (let iter = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()); iter <= horizonDate; iter = addDays(iter, 7)) {
      const iso = isoFromDate(iter);
      const serviceId = `${iso}_10am`;
      if (existing.has(serviceId)) continue;
      existing.add(serviceId);

      const row = Array.from({ length: lastCol }, () => '');
      row[idIdx] = serviceId;
      if (dateIdx >= 0) row[dateIdx] = new Date(iter.getFullYear(), iter.getMonth(), iter.getDate());
      if (timeIdx >= 0) row[timeIdx] = DEFAULT_SERVICE_TIME;
      const svcType = defaultServiceTypeForDate(iter);
      if (typeIdx >= 0) row[typeIdx] = svcType;
      if (leaderIdx >= 0) row[leaderIdx] = DEFAULT_LEADER;
      if (preacherIdx >= 0) row[preacherIdx] = DEFAULT_PREACHER;
      if (suggestedSongsIdx >= 0) row[suggestedSongsIdx] = '';
      rows.push(row);
      created.push({ id: serviceId, date: iso, time: DEFAULT_SERVICE_TIME, type: svcType });
    }

    if (rows.length) {
      const startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, rows.length, lastCol).setValues(rows);
      try { CacheService.getDocumentCache().remove(SERVICES_CACHE_KEY); } catch (_) {}
    }

    return { created };
  } finally {
    lock.releaseLock();
  }
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
  const suggestedSongsIdx = col(SERVICES_COL.suggestedSongs);

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
      if (suggestedSongsIdx >= 0) row[suggestedSongsIdx] = input.suggestedSongs ?? '';
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

  let lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  lastCol = ensureServiceColumns(sh, headers, [SERVICES_COL.youtubeUrl, SERVICES_COL.suggestedSongs]);
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const idIdx = col(SERVICES_COL.id);
  const typeIdx = col(SERVICES_COL.type);
  const youtubeUrlIdx = col(SERVICES_COL.youtubeUrl);
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
  const suggestedSongsIdx = col(SERVICES_COL.suggestedSongs);

  const originalId = String(input.id || '').trim();
  if (!originalId) throw new Error('Service ID is required to update a service.');
  const newId = originalId;

  // Build row data according to headers
  const vals: any[] = Array.from({ length: lastCol }, () => '');
  if (idIdx >= 0) vals[idIdx] = newId;
  if (typeIdx >= 0) vals[typeIdx] = input.type ?? '';
  if (youtubeUrlIdx >= 0) vals[youtubeUrlIdx] = input.youtubeUrl ?? '';
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
  if (suggestedSongsIdx >= 0) vals[suggestedSongsIdx] = input.suggestedSongs ?? '';

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

export function getSongPerformances(input: { songName?: string } | string) {
  const songName = typeof input === 'string'
    ? String(input || '').trim()
    : String((input as any)?.songName || '').trim();
  if (!songName) return { items: [] as SongPerformance[] };

  const target = normalizeSongLookup(songName);
  if (!target) return { items: [] as SongPerformance[] };

  const services = fetchServicesUnfiltered();
  const streamUrlByServiceId = getYouTubeStreamUrlByServiceIdMap();
  const performanceLinkByServiceId = getSongPerformanceLinkMap(songName);
  const serviceById = new Map<string, ServiceItem>();
  services.forEach(service => {
    const id = String(service?.id || '').trim();
    if (id) serviceById.set(id, service);
  });

  const sh = getSheetByName(ORDER_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { items: [] as SongPerformance[] };

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const serviceIdx = col(ORDER_COL.serviceId);
  const typeIdx = col(ORDER_COL.itemType);
  const detailIdx = col(ORDER_COL.detail);
  if (serviceIdx < 0 || typeIdx < 0 || detailIdx < 0) return { items: [] as SongPerformance[] };

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const seen = new Set<string>();
  const items: SongPerformance[] = [];
  const cutoff = todayISO();
  for (const row of body) {
    const itemType = String(row[typeIdx] ?? '').trim();
    if (!looksLikeSongItemType(itemType)) continue;
    const detail = String(row[detailIdx] ?? '').trim();
    if (!detail || normalizeSongLookup(detail) !== target) continue;
    const serviceId = String(row[serviceIdx] ?? '').trim();
    if (!serviceId || seen.has(serviceId)) continue;
    seen.add(serviceId);
    const service = serviceById.get(serviceId);
    const date = String(service?.date || deriveDateFromServiceId(serviceId) || '').trim();
    if (date && date > cutoff) continue;
    const time = String(service?.time || deriveTimeFromServiceId(serviceId) || '').trim();
    const type = String(service?.type || '').trim();
    const linkRow = performanceLinkByServiceId.get(serviceId);
    const baseYoutubeUrl = String(linkRow?.youtubeUrl || service?.youtubeUrl || streamUrlByServiceId.get(serviceId) || '').trim();
    const startSeconds = Math.max(0, Math.floor(Number(linkRow?.startSeconds) || 0));
    items.push({
      serviceId,
      date,
      time,
      type,
      label: formatPerformanceLabel({ date, time, type }),
      youtubeUrl: appendYouTubeStartTime(baseYoutubeUrl, startSeconds),
      baseYoutubeUrl,
      startSeconds: startSeconds || undefined,
      startLabel: String(linkRow?.startLabel || (startSeconds > 0 ? formatSecondsAsTimestamp(startSeconds) : '')).trim() || undefined
    });
  }

  items.sort((a, b) => serviceSortKey(b as ServiceItem).localeCompare(serviceSortKey(a as ServiceItem)));
  return { items };
}

export function saveSongPerformanceTimestamp(input?: SaveSongPerformanceTimestampInput) {
  const songName = String(input?.songName || '').trim();
  const serviceId = String(input?.serviceId || '').trim();
  const explicitUrl = String(input?.youtubeUrl || '').trim();
  const timestampInput = input?.timestampInput;

  if (!songName) throw new Error('Song name is required.');
  if (!serviceId) throw new Error('Service ID is required.');

  const startSeconds = parseTimestampToSeconds(timestampInput);
  if (!startSeconds) {
    throw new Error('Paste a YouTube link copied at the current time, or enter a timestamp like 17:36.');
  }

  const services = fetchServicesUnfiltered();
  const service = services.find(item => String(item?.id || '').trim() === serviceId);
  const fallbackUrl = String(service?.youtubeUrl || getYouTubeStreamUrlByServiceIdMap().get(serviceId) || '').trim();
  const youtubeUrl = explicitUrl || fallbackUrl;
  if (!youtubeUrl) {
    throw new Error('No YouTube recording is linked for this service yet.');
  }

  const headers = Object.values(SONG_PERFORMANCES_COL);
  const sh = getOrCreateSheet(SONG_PERFORMANCES_SHEET, headers);
  const lastCol = sh.getLastColumn();
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  ensureServiceColumns(sh, headerRow, headers);
  const normalizedHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => normalizedHeaders.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const songIdIdx = col(SONG_PERFORMANCES_COL.songId);
  const songNameIdx = col(SONG_PERFORMANCES_COL.songName);
  const serviceIdIdx = col(SONG_PERFORMANCES_COL.serviceId);
  const youtubeUrlIdx = col(SONG_PERFORMANCES_COL.youtubeUrl);
  const videoIdIdx = col(SONG_PERFORMANCES_COL.videoId);
  const startSecondsIdx = col(SONG_PERFORMANCES_COL.startSeconds);
  const startLabelIdx = col(SONG_PERFORMANCES_COL.startLabel);
  const matchSourceIdx = col(SONG_PERFORMANCES_COL.matchSource);
  const matchConfidenceIdx = col(SONG_PERFORMANCES_COL.matchConfidence);
  const lastVerifiedIdx = col(SONG_PERFORMANCES_COL.lastVerified);

  const nextRow = Array.from({ length: sh.getLastColumn() }, () => '');
  const normalizedSongId = normalizeSongLookup(songName);
  if (songIdIdx >= 0) nextRow[songIdIdx] = normalizedSongId;
  if (songNameIdx >= 0) nextRow[songNameIdx] = songName;
  if (serviceIdIdx >= 0) nextRow[serviceIdIdx] = serviceId;
  if (youtubeUrlIdx >= 0) nextRow[youtubeUrlIdx] = youtubeUrl;
  if (videoIdIdx >= 0) nextRow[videoIdIdx] = extractVideoIdFromYouTubeUrl(youtubeUrl);
  if (startSecondsIdx >= 0) nextRow[startSecondsIdx] = startSeconds;
  if (startLabelIdx >= 0) nextRow[startLabelIdx] = formatSecondsAsTimestamp(startSeconds);
  if (matchSourceIdx >= 0) nextRow[matchSourceIdx] = 'manual';
  if (matchConfidenceIdx >= 0) nextRow[matchConfidenceIdx] = 'verified';
  if (lastVerifiedIdx >= 0) nextRow[lastVerifiedIdx] = new Date();

  const lastRow = sh.getLastRow();
  let rowOffset = -1;
  if (lastRow >= 2 && serviceIdIdx >= 0) {
    const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    rowOffset = rows.findIndex(row =>
      normalizeSongLookup(String(row[songNameIdx] ?? row[songIdIdx] ?? '')) === normalizedSongId &&
      String(row[serviceIdIdx] ?? '').trim() === serviceId
    );
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    if (rowOffset >= 0) {
      sh.getRange(rowOffset + 2, 1, 1, nextRow.length).setValues([nextRow]);
    } else {
      sh.appendRow(nextRow);
    }
  } finally {
    lock.releaseLock();
  }

  return {
    ok: true,
    serviceId,
    songName,
    startSeconds,
    startLabel: formatSecondsAsTimestamp(startSeconds),
    youtubeUrl: appendYouTubeStartTime(youtubeUrl, startSeconds)
  };
}

export function suggestYouTubeStream(input: { date?: string }): SuggestYouTubeStreamResult {
  const isoDate = normalizeIso(input?.date || '') || '';
  if (!isoDate) {
    return {
      url: '',
      title: '',
      matchType: 'none',
      message: 'Choose a valid service date first.',
      candidates: []
    };
  }

  const scored = collectScoredYouTubeCandidates(isoDate, true);
  if (!scored.length) {
    return {
      url: '',
      title: '',
      matchType: 'none',
      message: `Could not load recent YouTube entries. Open ${DEFAULT_YOUTUBE_STREAMS_URL} to pick the stream manually.`,
      candidates: []
    };
  }

  const top = scored.slice(0, 5).map(item => ({
    title: item.title,
    url: item.url,
    published: item.published,
    reason: item.reason
  }));
  const best = scored[0];
  if (!best) {
    return {
      url: '',
      title: '',
      matchType: 'none',
      message: `No likely stream match found for ${formatIsoDateLabel(isoDate)}. Open the channel streams page and paste the correct link manually.`,
      candidates: []
    };
  }

  const matchType: SuggestYouTubeStreamResult['matchType'] =
    best.score >= 100 ? 'exact-title' :
    best.score >= 72 ? 'same-day' :
    'nearby';
  const message =
    matchType === 'exact-title'
      ? `Found an exact title match for ${formatIsoDateLabel(isoDate)}.`
      : matchType === 'same-day'
        ? `Found a same-day YouTube entry for ${formatIsoDateLabel(isoDate)}. Please confirm it is the right stream.`
        : `Found a nearby YouTube entry for ${formatIsoDateLabel(isoDate)}. Please confirm it before saving.`;

  return {
    url: best.url,
    title: best.title,
    matchType,
    message,
    candidates: top
  };
}

export function debugYouTubeMatching(options?: { limit?: number }) {
  const limit = Math.max(1, Math.min(25, Number(options?.limit || 10)));
  const services = fetchServicesUnfiltered()
    .filter(service => !String(service.youtubeUrl || '').trim())
    .filter(service => {
      const date = canonicalServiceDate(String(service.id || '').trim(), service.date);
      return !!date && date <= todayISO();
    })
    .sort((a, b) => serviceSortKey(b).localeCompare(serviceSortKey(a)))
    .slice(0, limit);

  const ss = SpreadsheetApp.getActive();
  const name = 'YouTube Debug';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();

  const rows: any[][] = [[
    'ServiceID', 'ServiceDate', 'CandidateRank', 'CandidateTitle', 'CandidateUrl', 'CandidatePublished', 'Score', 'Reason'
  ]];

  services.forEach(service => {
    const date = canonicalServiceDate(String(service.id || '').trim(), service.date);
    const scored = collectScoredYouTubeCandidates(date, true).slice(0, 5);
    if (!scored.length) {
      rows.push([service.id, date, '', '(no candidates)', '', '', '', 'No candidates found']);
      return;
    }
    scored.forEach((cand, index) => {
      rows.push([
        service.id,
        date,
        index + 1,
        cand.title,
        cand.url,
        cand.published,
        cand.score,
        cand.reason
      ]);
    });
  });

  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  try { sh.autoResizeColumns(1, rows[0].length); } catch (_) {}
  try { SpreadsheetApp.getActive().toast(`YouTube debug report written to "${name}"`, 'Worship Planner', 5); } catch (_) {}
  return { sheet: name, services: services.length, rows: rows.length - 1 };
}

export function debugYouTubeFetch() {
  const channelId = getYouTubeChannelId();
  const targets = [
    DEFAULT_YOUTUBE_CHANNEL_URL,
    DEFAULT_YOUTUBE_STREAMS_URL,
    channelId ? `https://www.youtube.com/feeds/videos.xml?channel_id=${encodeURIComponent(channelId)}` : '',
    'https://www.youtube.com/results?search_query=' + encodeURIComponent(`${DEFAULT_YOUTUBE_TITLE_PREFIX} 19 July 2026`)
  ].filter(Boolean);

  const ss = SpreadsheetApp.getActive();
  const name = 'YouTube Fetch Debug';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();

  const rows: any[][] = [[
    'Url', 'Status', 'Length', 'HasWatchHref', 'HasConsent', 'HasChannelId', 'HasDatePublished', 'HasStreamedLive', 'TitleSample', 'Snippet'
  ]];

  targets.forEach((target) => {
    const info = fetchYouTubeUrlDebug(target);
    rows.push([
      info.url,
      info.status,
      info.length,
      info.hasWatchHref ? 'Y' : '',
      info.hasConsent ? 'Y' : '',
      info.hasChannelId ? 'Y' : '',
      info.hasDatePublished ? 'Y' : '',
      info.hasStreamedLive ? 'Y' : '',
      info.titleSample,
      info.snippet
    ]);
  });

  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  try { sh.autoResizeColumns(1, rows[0].length); } catch (_) {}
  try { SpreadsheetApp.getActive().toast(`YouTube fetch debug written to "${name}"`, 'Worship Planner', 5); } catch (_) {}
  return { sheet: name, rows: rows.length - 1, channelId };
}

export function debugYouTubeExtraction() {
  const ss = SpreadsheetApp.getActive();
  const name = 'YouTube Extract Debug';
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();

  const searchQuery = `${DEFAULT_YOUTUBE_TITLE_PREFIX} 19 July 2026`;
  const streamsResponse = UrlFetchApp.fetch(DEFAULT_YOUTUBE_STREAMS_URL, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Google-Apps-Script)' },
    muteHttpExceptions: true
  });
  const searchResponse = UrlFetchApp.fetch(`https://www.youtube.com/results?search_query=${encodeURIComponent(searchQuery)}`, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Google-Apps-Script)' },
    muteHttpExceptions: true
  });

  const streamsHtml = streamsResponse.getContentText();
  const searchHtml = searchResponse.getContentText();
  const streamsItems = extractYouTubeEntriesFromHtml(streamsHtml);
  const searchItems = extractYouTubeEntriesFromHtml(searchHtml);

  const markerInfo = (html: string) => {
    const text = String(html || '');
    const inspected = inspectYouTubeInitialData(text);
    const anchorCount = extractYouTubeEntriesViaAnchors(text).length;
    const watchCount = extractYouTubeEntriesViaWatchSnippets(text).length;
    return {
      hasVarMarker: text.includes('var ytInitialData = '),
      hasWindowMarker: text.includes('window["ytInitialData"] = '),
      hasPlainMarker: text.includes('ytInitialData = '),
      hasJsonParse: text.includes('JSON.parse('),
      loaded: !!inspected.value,
      parseMarker: inspected.parseMarker || inspected.rawMarker,
      parsedStringLength: inspected.parsedStringLength,
      quotedStringLength: inspected.quotedStringLength,
      rawObjectLength: inspected.rawObjectLength,
      parseError: inspected.parseError,
      markerSnippet: inspected.snippet,
      anchorCount,
      watchCount
    };
  };
  const streamsInfo = markerInfo(streamsHtml);
  const searchInfo = markerInfo(searchHtml);

  const rows: any[][] = [[
    'Source', 'Status', 'ExtractedCount', 'AnchorCount', 'WatchCount', 'HasVarMarker', 'HasWindowMarker', 'HasPlainMarker', 'HasJsonParse', 'LoadedInitialData', 'ParseMarker', 'ParsedStringLength', 'QuotedStringLength', 'RawObjectLength', 'ParseError', 'MarkerSnippet', 'Title', 'Url', 'Published'
  ]];

  rows.push(['streams', streamsResponse.getResponseCode(), streamsItems.length, streamsInfo.anchorCount, streamsInfo.watchCount, streamsInfo.hasVarMarker ? 'Y' : '', streamsInfo.hasWindowMarker ? 'Y' : '', streamsInfo.hasPlainMarker ? 'Y' : '', streamsInfo.hasJsonParse ? 'Y' : '', streamsInfo.loaded ? 'Y' : '', streamsInfo.parseMarker, streamsInfo.parsedStringLength, streamsInfo.quotedStringLength, streamsInfo.rawObjectLength, streamsInfo.parseError, streamsInfo.markerSnippet, '', '', '']);
  streamsItems.slice(0, 25).forEach(item => {
    rows.push(['streams', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', item.title, item.url, item.published]);
  });

  rows.push(['search', searchResponse.getResponseCode(), searchItems.length, searchInfo.anchorCount, searchInfo.watchCount, searchInfo.hasVarMarker ? 'Y' : '', searchInfo.hasWindowMarker ? 'Y' : '', searchInfo.hasPlainMarker ? 'Y' : '', searchInfo.hasJsonParse ? 'Y' : '', searchInfo.loaded ? 'Y' : '', searchInfo.parseMarker, searchInfo.parsedStringLength, searchInfo.quotedStringLength, searchInfo.rawObjectLength, searchInfo.parseError, searchInfo.markerSnippet, '', '', '']);
  searchItems.slice(0, 25).forEach(item => {
    rows.push(['search', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', item.title, item.url, item.published]);
  });

  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  try { sh.autoResizeColumns(1, rows[0].length); } catch (_) {}
  try { SpreadsheetApp.getActive().toast(`YouTube extraction debug written to "${name}"`, 'Worship Planner', 5); } catch (_) {}
  return { sheet: name, streams: streamsItems.length, search: searchItems.length };
}

export function syncYouTubeStreamsCatalog() {
  const minDate = defaultYouTubeBackfillStartDate();
  const cursor = getYouTubeCatalogCursor();
  const headers = [
    YOUTUBE_STREAMS_COL.videoId,
    YOUTUBE_STREAMS_COL.url,
    YOUTUBE_STREAMS_COL.title,
    YOUTUBE_STREAMS_COL.streamDate,
    YOUTUBE_STREAMS_COL.publishedDate,
    YOUTUBE_STREAMS_COL.channelId,
    YOUTUBE_STREAMS_COL.channelName,
    YOUTUBE_STREAMS_COL.matchedServiceId,
    YOUTUBE_STREAMS_COL.status,
    YOUTUBE_STREAMS_COL.notes,
    YOUTUBE_STREAMS_COL.source
  ];
  const sh = getOrCreateSheet(YOUTUBE_STREAMS_SHEET, headers);
  const existingLastRow = sh.getLastRow();
  const existingLastCol = Math.max(sh.getLastColumn(), headers.length);
  const existingRows = existingLastRow > 1
    ? sh.getRange(2, 1, existingLastRow - 1, existingLastCol).getValues()
    : [];
  const existingByVideoId = new Map<string, { row: any[]; rowIndex: number }>();
  existingRows.forEach((row, index) => {
    const videoId = String(row[0] ?? '').trim();
    if (videoId) existingByVideoId.set(videoId, { row, rowIndex: index + 2 });
  });

  const result = listYouTubeStreamCatalogEntries({ pageToken: cursor });
  const rows = result.items.map((entry) => {
    const existing = existingByVideoId.get(entry.videoId)?.row || [];
    const matchedServiceId = String(existing[7] ?? '').trim();
    const existingNotes = String(existing[9] ?? '').trim();
    const status = matchedServiceId
      ? 'Matched'
      : entry.streamDate
        ? 'Ready'
        : 'Needs Review';
    return [
      entry.videoId,
      entry.url,
      entry.title,
      entry.streamDate ? toSheetDateValue(entry.streamDate) : '',
      entry.published ? toSheetDateValue(entry.published) : '',
      entry.channelId,
      entry.channelName,
      matchedServiceId,
      status,
      existingNotes,
      entry.source
    ];
  });
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  const mergedRows = new Map<string, any[]>();
  existingRows.forEach((row) => {
    const videoId = String(row[0] ?? '').trim();
    if (!videoId) return;
    const normalized = row.slice(0, headers.length);
    const streamIso = normalizeIso(normalized[3] as any) || '';
    const publishedIso = normalizeIso(normalized[4] as any) || '';
    normalized[3] = streamIso ? toSheetDateValue(streamIso) : '';
    normalized[4] = publishedIso ? toSheetDateValue(publishedIso) : '';
    mergedRows.set(videoId, normalized);
  });
  rows.forEach((row) => {
    const videoId = String(row[0] ?? '').trim();
    if (videoId) mergedRows.set(videoId, row);
  });
  const finalRows = Array.from(mergedRows.values()).sort((a, b) => {
    const aDate = normalizeIso(a[3] as any) || '';
    const bDate = normalizeIso(b[3] as any) || '';
    return `${bDate}|${String(b[2] || '')}`.localeCompare(`${aDate}|${String(a[2] || '')}`);
  });
  if (finalRows.length) sh.getRange(2, 1, finalRows.length, headers.length).setValues(finalRows);
  const done = !!result.reachedEnd;
  if (done) clearYouTubeCatalogCursor();
  else setYouTubeCatalogCursor(String(result.nextPageToken || ''));
  try { sh.autoResizeColumns(1, headers.length); } catch (_) {}
  try {
    SpreadsheetApp.getActive().toast(
      done
        ? `YouTube streams catalog complete from ${minDate}: ${finalRows.length} row${finalRows.length === 1 ? '' : 's'}`
        : `YouTube streams catalog progress: processed ${result.processedEntries} upload${result.processedEntries === 1 ? '' : 's'} from playlist`,
      'Worship Planner',
      5
    );
  } catch (_) {}
  return {
    sheet: YOUTUBE_STREAMS_SHEET,
    rowsUpdated: rows.length,
    processedEntries: result.processedEntries,
    totalEntries: result.totalEntries,
    nextPageToken: result.nextPageToken,
    done
  };
}

export function resetYouTubeStreamsSyncState() {
  clearYouTubeCatalogCursor();
  try {
    SpreadsheetApp.getActive().toast('YouTube streams sync state reset.', 'Worship Planner', 5);
  } catch (_) {}
  return { ok: true };
}

export function matchServicesFromYouTubeStreams(options?: { overwriteExisting?: boolean; limit?: number }) {
  const minDate = defaultYouTubeBackfillStartDate();
  const servicesSheet = getSheetByName(SERVICES_SHEET);
  let lastCol = servicesSheet.getLastColumn();
  const serviceHeaders = servicesSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  lastCol = ensureServiceColumns(servicesSheet, serviceHeaders, [SERVICES_COL.youtubeUrl]);
  const serviceCol = (name: string) => serviceHeaders.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const serviceIdIdx = serviceCol(SERVICES_COL.id);
  const serviceDateIdx = serviceCol(SERVICES_COL.date);
  const serviceYoutubeIdx = serviceCol(SERVICES_COL.youtubeUrl);
  if (serviceIdIdx < 0 || serviceDateIdx < 0 || serviceYoutubeIdx < 0) {
    throw new Error('Services sheet is missing required columns for YouTube matching.');
  }

  const streamsHeaders = [
    YOUTUBE_STREAMS_COL.videoId,
    YOUTUBE_STREAMS_COL.url,
    YOUTUBE_STREAMS_COL.title,
    YOUTUBE_STREAMS_COL.streamDate,
    YOUTUBE_STREAMS_COL.publishedDate,
    YOUTUBE_STREAMS_COL.channelId,
    YOUTUBE_STREAMS_COL.channelName,
    YOUTUBE_STREAMS_COL.matchedServiceId,
    YOUTUBE_STREAMS_COL.status,
    YOUTUBE_STREAMS_COL.notes,
    YOUTUBE_STREAMS_COL.source
  ];
  const streamsSheet = getOrCreateSheet(YOUTUBE_STREAMS_SHEET, streamsHeaders);
  const streamsLastRow = streamsSheet.getLastRow();
  const streamsBody = streamsLastRow > 1
    ? streamsSheet.getRange(2, 1, streamsLastRow - 1, streamsHeaders.length).getValues()
    : [];
  const byDate = new Map<string, Array<{ rowIndex: number; row: any[] }>>();
  streamsBody.forEach((row, index) => {
    const streamDate = normalizeIso(row[3] as any) || '';
    const url = String(row[1] ?? '').trim();
    if (!streamDate || !url) return;
    const arr = byDate.get(streamDate) || [];
    arr.push({ rowIndex: index + 2, row });
    byDate.set(streamDate, arr);
  });

  const reviewSheetName = 'YouTube Stream Match Review';
  const reviewSheet = getOrCreateSheet(reviewSheetName, [
    'ServiceID', 'ServiceDate', 'Status', 'CandidateCount', 'CandidateTitles', 'CandidateUrls'
  ]);
  reviewSheet.clearContents();
  reviewSheet.getRange(1, 1, 1, 6).setValues([['ServiceID', 'ServiceDate', 'Status', 'CandidateCount', 'CandidateTitles', 'CandidateUrls']]);
  const reviewRows: any[][] = [];

  const limit = Math.max(0, Number(options?.limit || 0)) || 0;
  const overwriteExisting = !!options?.overwriteExisting;
  const cutoff = todayISO();
  const body = servicesSheet.getRange(2, 1, servicesSheet.getLastRow() - 1, lastCol).getValues();
  let matched = 0;
  let skipped = 0;

  for (let i = 0; i < body.length; i++) {
    if (limit > 0 && matched + skipped >= limit) break;
    const row = body[i];
    const serviceId = String(row[serviceIdIdx] ?? '').trim();
    const existingUrl = String(row[serviceYoutubeIdx] ?? '').trim();
    const date = canonicalServiceDate(serviceId, row[serviceDateIdx]);
    if (!date || date > cutoff || date < minDate) continue;

    const candidates = (byDate.get(date) || []).filter(item => {
      const matchedServiceId = String(item.row[7] ?? '').trim();
      return !matchedServiceId || matchedServiceId === serviceId;
    });

    if (candidates.length === 1) {
      const candidate = candidates[0];
      const url = String(candidate.row[1] ?? '').trim();
      if (!existingUrl || overwriteExisting) {
        servicesSheet.getRange(i + 2, serviceYoutubeIdx + 1).setValue(url);
      }
      streamsSheet.getRange(candidate.rowIndex, 8, 1, 2).setValues([[serviceId, 'Matched']]);
      matched += 1;
      continue;
    }

    skipped += 1;
    reviewRows.push([
      serviceId,
      date,
      candidates.length ? 'Ambiguous' : 'No stream found',
      candidates.length,
      candidates.map(item => String(item.row[2] ?? '').trim()).join(' | '),
      candidates.map(item => String(item.row[1] ?? '').trim()).join(' | ')
    ]);
  }

  if (reviewRows.length) reviewSheet.getRange(2, 1, reviewRows.length, 6).setValues(reviewRows);
  try { reviewSheet.autoResizeColumns(1, 6); } catch (_) {}
  try { SpreadsheetApp.flush(); } catch (_) {}
  try { CacheService.getDocumentCache().remove(SERVICES_CACHE_KEY); } catch (_) {}
  try {
    SpreadsheetApp.getActive().toast(`YouTube stream match from ${minDate}: ${matched} matched, ${skipped} need review`, 'Worship Planner', 5);
  } catch (_) {}
  return { matched, skipped, reviewSheet: reviewSheetName, catalogSheet: YOUTUBE_STREAMS_SHEET };
}

export function syncMissingYouTubeUrls(options?: { includeNearby?: boolean; limit?: number }) {
  const sh = getSheetByName(SERVICES_SHEET);
  let lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  lastCol = ensureServiceColumns(sh, headers, [SERVICES_COL.youtubeUrl]);
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const idIdx = col(SERVICES_COL.id);
  const dateIdx = col(SERVICES_COL.date);
  const youtubeUrlIdx = col(SERVICES_COL.youtubeUrl);
  if (youtubeUrlIdx < 0) throw new Error(`Column "${SERVICES_COL.youtubeUrl}" not found in ${SERVICES_SHEET}`);
  if (lastCol < 1 || sh.getLastRow() < 2) return { updated: 0, skipped: 0, reviewed: 0, items: [] };

  const includeNearby = !!options?.includeNearby;
  const limit = Math.max(0, Number(options?.limit || 0)) || 0;
  const cutoff = todayISO();
  const body = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getValues();

  const toIso = (v: unknown) => {
    try {
      if (v instanceof Date && !isNaN(v.getTime())) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, '0');
        const d = String(v.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
    } catch (_) {}
    const raw = String(v ?? '').trim();
    if (ISO_DATE_RE.test(raw)) return raw;
    return '';
  };

  const reviewItems: Array<{ row: number; serviceId: string; date: string; matchType: string; title: string; url: string; message: string }> = [];
  let updated = 0;
  let skipped = 0;
  let processed = 0;
  const totalCandidates = body.filter((row) => {
    const serviceId = idIdx >= 0 ? String(row[idIdx] ?? '').trim() : '';
    const existingUrl = String(row[youtubeUrlIdx] ?? '').trim();
    if (existingUrl) return false;
    const date = canonicalServiceDate(serviceId, dateIdx >= 0 ? row[dateIdx] : '');
    return !!date && date <= cutoff;
  }).length;

  try {
    SpreadsheetApp.getActive().toast(`YouTube sync started for ${totalCandidates} service${totalCandidates === 1 ? '' : 's'}`, 'Worship Planner', 5);
  } catch (_) {}

  for (let i = 0; i < body.length; i++) {
    if (limit > 0 && updated + skipped >= limit) break;
    const row = body[i];
    const serviceId = idIdx >= 0 ? String(row[idIdx] ?? '').trim() : '';
    const existingUrl = String(row[youtubeUrlIdx] ?? '').trim();
    if (existingUrl) continue;

    const date = canonicalServiceDate(serviceId, dateIdx >= 0 ? row[dateIdx] : '');
    if (!date || date > cutoff) continue;
    processed += 1;
    if (processed === 1 || processed % 5 === 0) {
      try {
        SpreadsheetApp.getActive().toast(`YouTube sync: matched ${updated}, checked ${processed}/${totalCandidates}`, 'Worship Planner', 3);
      } catch (_) {}
    }

    const result = suggestYouTubeStream({ date });
    if (result.matchType === 'exact-title' || result.matchType === 'same-day' || (includeNearby && result.matchType === 'nearby')) {
      const nextUrl = String(result.url || '').trim();
      if (nextUrl) {
        sh.getRange(i + 2, youtubeUrlIdx + 1).setValue(nextUrl);
        try { SpreadsheetApp.flush(); } catch (_) {}
        updated += 1;
        reviewItems.push({
          row: i + 2,
          serviceId,
          date,
          matchType: result.matchType,
          title: String(result.title || '').trim(),
          url: nextUrl,
          message: result.message
        });
        continue;
      }
    }

    skipped += 1;
  }

  try { CacheService.getDocumentCache().remove(SERVICES_CACHE_KEY); } catch (_) {}
  try {
    const summary = `YouTube sync: ${updated} matched, ${skipped} still need review`;
    SpreadsheetApp.getActive().toast(summary, 'Worship Planner', 5);
  } catch (_) {}

  return {
    updated,
    skipped,
    reviewed: reviewItems.length,
    items: reviewItems
  };
}

export function repairServiceDateTimeColumns() {
  const sh = getSheetByName(SERVICES_SHEET);
  let lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  lastCol = ensureServiceColumns(sh, headers, [SERVICES_COL.date, SERVICES_COL.time, SERVICES_COL.type]);
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const idIdx = col(SERVICES_COL.id);
  const dateIdx = col(SERVICES_COL.date);
  const timeIdx = col(SERVICES_COL.time);
  const typeIdx = col(SERVICES_COL.type);
  if (idIdx < 0 || dateIdx < 0 || timeIdx < 0 || typeIdx < 0) {
    throw new Error(`Required columns missing in ${SERVICES_SHEET}`);
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { updated: 0 };
  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let updated = 0;

  for (let i = 0; i < body.length; i++) {
    const row = body[i];
    const serviceId = String(row[idIdx] ?? '').trim();
    if (!serviceId) continue;
    const derivedDate = deriveDateFromServiceId(serviceId);
    const derivedTime = deriveTimeFromServiceId(serviceId);

    let touched = false;
    if (derivedDate) {
      const normalizedCurrentDate = normalizeIso(row[dateIdx] as any) || '';
      if (normalizedCurrentDate !== derivedDate) {
        sh.getRange(i + 2, dateIdx + 1).setValue(toSheetDateValue(derivedDate));
        touched = true;
      }
    }

    if (derivedTime) {
      const currentTime = String(row[timeIdx] ?? '').trim();
      if (!currentTime || currentTime !== derivedTime) {
        sh.getRange(i + 2, timeIdx + 1).setValue(derivedTime);
        touched = true;
      }
    }

    const currentType = String(row[typeIdx] ?? '').trim();
    if (!currentType && derivedDate) {
      const derivedDateObj = dateFromISO(derivedDate);
      if (derivedDateObj) {
        sh.getRange(i + 2, typeIdx + 1).setValue(defaultServiceTypeForDate(derivedDateObj));
        touched = true;
      }
    }

    if (touched) updated += 1;
  }

  try { CacheService.getDocumentCache().remove(SERVICES_CACHE_KEY); } catch (_) {}
  try {
    SpreadsheetApp.getActive().toast(`Service repair complete: ${updated} row${updated === 1 ? '' : 's'} updated`, 'Worship Planner', 5);
  } catch (_) {}
  return { updated };
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
  return readDocumentCachedJson<{ leaders: string[]; preachers: string[] }>({
    key: SERVICE_PEOPLE_CACHE_KEY,
    ttlSeconds: 300,
    version: getSpreadsheetVersion([SERVICES_SHEET, PLANNER_SHEET]),
    loader: () => {
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
  });
}

const escapeRegex = (value: string) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const normalizeReferenceSpacing = (value: string) => String(value || '').replace(/\s+/g, ' ').trim();

const BIBLE_BOOK_NAMES = [
  'Genesis','Exodus','Leviticus','Numbers','Deuteronomy',
  'Joshua','Judges','Ruth','1 Samuel','2 Samuel',
  '1 Kings','2 Kings','1 Chronicles','2 Chronicles','Ezra',
  'Nehemiah','Esther','Job','Psalms','Proverbs',
  'Ecclesiastes','Song of Solomon','Isaiah','Jeremiah','Lamentations',
  'Ezekiel','Daniel','Hosea','Joel','Amos','Obadiah',
  'Jonah','Micah','Nahum','Habakkuk','Zephaniah','Haggai',
  'Zechariah','Malachi',
  'Matthew','Mark','Luke','John','Acts',
  'Romans','1 Corinthians','2 Corinthians','Galatians','Ephesians',
  'Philippians','Colossians','1 Thessalonians','2 Thessalonians','1 Timothy',
  '2 Timothy','Titus','Philemon','Hebrews','James',
  '1 Peter','2 Peter','1 John','2 John','3 John','Jude','Revelation'
];

const BOOK_REGEX_SRC = BIBLE_BOOK_NAMES
  .slice()
  .sort((a, b) => b.length - a.length)
  .map(name => escapeRegex(name).replace(/\s+/g, '\\s+'))
  .join('|');
const BOOK_REGEX_BODY = `(?:${BOOK_REGEX_SRC})`;
const BOOK_REGEX_WITH_BOUNDARY = `\\b${BOOK_REGEX_BODY}\\b`;

const REF_EDGE_PATTERN = '(?:[,;|/&+\\-\\u2013\\u2014]+|\\band\\b|&)';

const cleanPassageText = (input?: string) => {
  if (!input) return '';
  let text = String(input);
  // Remove bracketed footnote remnants or stray markers just in case
  text = text.replace(/\s*\[\d+\]\s*/g, ' ');
  // Normalize whitespace: trim, collapse 3+ newlines to 2, normalize CRLF
  text = text.replace(/\r\n?/g, '\n');
  return text.replace(/\n{3,}/g, '\n\n').trim();
};

const cleanPassageHtml = (input?: string) => {
  if (!input) return '';
  let html = String(input);
  // Basic cleanup: remove outer wrappers if present
  html = html.replace(/<p class=".*?">/g, '<p>').replace(/<h\d[^>]*>.*?<\/h\d>/g, '');
  return html.trim();
};

const trimReferenceConnectors = (segment: string) => {
  if (!segment) return '';
  let value = segment.trim();
  value = value.replace(new RegExp(`^${REF_EDGE_PATTERN}+\\s*`, 'i'), '').trim();
  value = value.replace(new RegExp(`\\s*${REF_EDGE_PATTERN}+$`, 'i'), '').trim();
  return value;
};

const splitReferenceIntoDistinctBooks = (reference: string): string[] => {
  const raw = String(reference || '');
  const regex = new RegExp(BOOK_REGEX_WITH_BOUNDARY, 'gi');
  const matches: Array<{ index: number; name: string }> = [];
  let match: RegExpExecArray | null;
  while ((match = regex.exec(raw)) !== null) {
    matches.push({ index: match.index, name: normalizeReferenceSpacing(match[0]) });
  }
  if (matches.length < 2) return [];
  const segments: Array<{ ref: string; book: string }> = [];
  for (let i = 0; i < matches.length; i += 1) {
    const start = matches[i].index;
    const end = i + 1 < matches.length ? matches[i + 1].index : raw.length;
    let chunk = raw.slice(start, end).trim();
    chunk = trimReferenceConnectors(chunk);
    if (!chunk) continue;
    segments.push({ ref: normalizeReferenceSpacing(chunk), book: matches[i].name.toLowerCase() });
  }
  const uniqueBooks = new Set(segments.map(s => s.book));
  if (segments.length >= 2 && uniqueBooks.size >= 2) {
    return segments.map(s => s.ref);
  }
  return [];
};

const escapeHtml = (input?: string) => {
  const str = String(input ?? '');
  return str.replace(/[&<>"']/g, (c) => {
    switch (c) {
      case '&': return '&amp;';
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '"': return '&quot;';
      case '\'': return '&#39;';
      default: return c;
    }
  });
};

type PassageChunk = { reference: string; text: string; html: string };
type PassageResult = { reference: string; text: string; html?: string; error?: string };

const fetchPassageChunk = (
  reference: string,
  token: string,
  includeHtml: boolean,
  includeInlineReference: boolean
): PassageChunk => {
  const normalizedRef = normalizeReferenceSpacing(reference);
  const textUrl = 'https://api.esv.org/v3/passage/text/?' +
    'q=' + encodeURIComponent(normalizedRef) +
    '&include-passage-references=false' +
    '&include-footnotes=false' +
    '&include-headings=false' +
    '&include-short-copyright=false' +
    '&include-verse-numbers=false' +
    '&indent-poetry=false' +
    '&indent-using=spaces' +
    '&indent-paragraphs=0';

  const res = UrlFetchApp.fetch(textUrl, { headers: { Authorization: 'Token ' + token } });
  const data = JSON.parse(res.getContentText());
  const rawPassages = Array.isArray(data?.passages) ? data.passages : [];
  const textParts = rawPassages.map(p => cleanPassageText(String(p ?? ''))).filter(Boolean);
  let text = textParts.join('\n\n').trim();
  if (includeInlineReference && normalizedRef && textParts.length > 1 && text) {
    text = `${normalizedRef}\n\n${text}`;
  }

  let html = '';
  if (includeHtml) {
    try {
      const htmlUrl = 'https://api.esv.org/v3/passage/html/?' +
        'q=' + encodeURIComponent(normalizedRef) +
        '&include-passage-references=false' +
        '&include-footnotes=false' +
        '&include-headings=false' +
        '&include-short-copyright=false' +
        '&include-verse-numbers=true' +
        '&inline-styles=false';
      const hres = UrlFetchApp.fetch(htmlUrl, { headers: { Authorization: 'Token ' + token } });
      const hdata = JSON.parse(hres.getContentText());
      const htmlPassages = Array.isArray(hdata?.passages) ? hdata.passages : [];
      html = htmlPassages.map(p => cleanPassageHtml(String(p ?? ''))).filter(Boolean).join('<hr />').trim();
    } catch (_) { /* ignore html errors */ }
  }

  return { reference: normalizedRef, text, html };
};

const formatChunkText = (reference: string, text: string) => {
  const cleanText = String(text || '').trim();
  if (!cleanText) return '';
  const ref = String(reference || '').trim();
  return ref ? `${ref}\n${cleanText}` : cleanText;
};

const formatChunkHtml = (reference: string, html: string, fallbackText: string) => {
  const body = String(html || '').trim() || (fallbackText ? `<p>${escapeHtml(fallbackText)}</p>` : '');
  if (!body) return '';
  const ref = String(reference || '').trim();
  const refBlock = ref ? `<p class="scripture-ref-block"><strong>${escapeHtml(ref)}</strong></p>` : '';
  return `<div class="scripture-chunk">${refBlock}${body}</div>`;
};

const decodeHtmlEntities = (input?: string) => {
  return String(input || '')
    .replace(/&#(\d+);/g, (_, dec) => {
      const code = Number(dec);
      return Number.isFinite(code) ? String.fromCharCode(code) : _;
    })
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => {
      const code = parseInt(hex, 16);
      return Number.isFinite(code) ? String.fromCharCode(code) : _;
    })
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, '\'')
    .replace(/&rsquo;/gi, '\'')
    .replace(/&lsquo;/gi, '\'')
    .replace(/&rdquo;/gi, '"')
    .replace(/&ldquo;/gi, '"')
    .replace(/&mdash;/gi, '-')
    .replace(/&ndash;/gi, '-')
    .replace(/&hellip;/gi, '...');
};

const stripBibleGatewayHtmlToText = (input?: string) => {
  let html = String(input || '');
  if (!html) return '';
  html = html
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<h[1-6]\b[^>]*>[\s\S]*?<\/h[1-6]>/gi, '')
    .replace(/<sup\b[^>]*>[\s\S]*?<\/sup>/gi, '')
    .replace(/<div\b[^>]*class="[^"]*(?:footnotes?|crossrefs?|crossreference)[^"]*"[^>]*>[\s\S]*?<\/div>/gi, '')
    .replace(/<div\b[^>]*class='[^']*(?:footnotes?|crossrefs?|crossreference)[^']*'[^>]*>[\s\S]*?<\/div>/gi, '')
    .replace(/<ol\b[^>]*class="[^"]*(?:footnotes?|crossrefs?|crossreference)[^"]*"[^>]*>[\s\S]*?<\/ol>/gi, '')
    .replace(/<ol\b[^>]*class='[^']*(?:footnotes?|crossrefs?|crossreference)[^']*'[^>]*>[\s\S]*?<\/ol>/gi, '')
    .replace(/<p\b[^>]*class="[^"]*(?:footnotes?|crossrefs?|crossreference)[^"]*"[^>]*>[\s\S]*?<\/p>/gi, '')
    .replace(/<p\b[^>]*class='[^']*(?:footnotes?|crossrefs?|crossreference)[^']*'[^>]*>[\s\S]*?<\/p>/gi, '')
    .replace(/<span\b[^>]*class="[^"]*(?:footnote|crossreference|chapternum|versenum)[^"]*"[^>]*>[\s\S]*?<\/span>/gi, '')
    .replace(/<span\b[^>]*class='[^']*(?:footnote|crossreference|chapternum|versenum)[^']*'[^>]*>[\s\S]*?<\/span>/gi, '')
    .replace(/<a\b[^>]*class="[^"]*full-chap-link[^"]*"[^>]*>[\s\S]*?<\/a>/gi, '')
    .replace(/<a\b[^>]*class='[^']*full-chap-link[^']*'[^>]*>[\s\S]*?<\/a>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<li\b[^>]*>/gi, '')
    .replace(/<[^>]+>/g, '');
  html = decodeHtmlEntities(html);
  html = html.replace(/\r\n?/g, '\n');
  html = html.replace(/[ \t]+\n/g, '\n');
  html = html.replace(/\n[ \t]+/g, '\n');
  html = html.replace(/\n{3,}/g, '\n\n');
  return html.trim();
};

const trimBibleGatewayAncillarySections = (input?: string) => {
  let html = String(input || '');
  if (!html) return '';
  html = html
    .replace(/<(?:div|section|ol|ul|p)\b[^>]*class="[^"]*(?:footnotes?|crossrefs?|crossreference)[^"]*"[^>]*>[\s\S]*$/i, '')
    .replace(/<(?:div|section|ol|ul|p)\b[^>]*class='[^']*(?:footnotes?|crossrefs?|crossreference)[^']*'[^>]*>[\s\S]*$/i, '')
    .replace(/<h[1-6]\b[^>]*>\s*(?:Footnotes|Cross references)\s*<\/h[1-6]>[\s\S]*$/i, '')
    .replace(/<div\b[^>]*id="(?:footnotes?|crossrefs?)"[^>]*>[\s\S]*$/i, '')
    .replace(/<div\b[^>]*id='(?:footnotes?|crossrefs?)'[^>]*>[\s\S]*$/i, '');
  return html.trim();
};

const extractBibleGatewayPassageHtml = (markup?: string) => {
  const html = String(markup || '');
  if (!html) return '';
  const patterns = [
    /<div\b[^>]*class="[^"]*passage-text[^"]*"[^>]*>([\s\S]*?)<div\b[^>]*class="[^"]*passage-meta[^"]*"[^>]*>/i,
    /<div\b[^>]*class="[^"]*passage-content[^"]*"[^>]*>([\s\S]*?)<\/div>\s*<\/div>\s*<div\b[^>]*class="[^"]*passage-meta[^"]*"[^>]*>/i,
    /<div\b[^>]*class="[^"]*passage-text[^"]*"[^>]*>([\s\S]*?)<\/article>/i,
    /<div\b[^>]*class="[^"]*passage-text[^"]*"[^>]*>([\s\S]*?)<a\b[^>]*class="[^"]*full-chap-link[^"]*"[^>]*>/i,
    /<div\b[^>]*class='[^']*passage-text[^']*'[^>]*>([\s\S]*?)<a\b[^>]*class='[^']*full-chap-link[^']*'[^>]*>/i,
    /<div\b[^>]*class="[^"]*passage-text[^"]*"[^>]*>([\s\S]*?)<div\b[^>]*class="[^"]*crossrefs[^"]*"[^>]*>/i,
    /<div\b[^>]*class='[^']*passage-text[^']*'[^>]*>([\s\S]*?)<div\b[^>]*class='[^']*crossrefs[^']*'[^>]*>/i
  ];
  for (const pattern of patterns) {
    const match = html.match(pattern);
    if (match && match[1]) return trimBibleGatewayAncillarySections(match[1]);
  }
  return '';
};

const fetchBibleGatewayChunk = (reference: string, version: string): PassageChunk => {
  const normalizedRef = normalizeReferenceSpacing(reference);
  const url = 'https://www.biblegateway.com/passage/?search=' +
    encodeURIComponent(normalizedRef) +
    '&version=' + encodeURIComponent(version) +
    '&interface=print';
  const res = UrlFetchApp.fetch(url, {
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Google-Apps-Script)' },
    muteHttpExceptions: true
  });
  const status = res.getResponseCode();
  if (status >= 400) throw new Error(`BibleGateway request failed (${status})`);
  const markup = res.getContentText();
  const passageHtml = extractBibleGatewayPassageHtml(markup);
  const text = stripBibleGatewayHtmlToText(passageHtml);
  if (!text) throw new Error('Unable to parse BibleGateway passage text');
  return { reference: normalizedRef, text, html: '' };
};

export function esvPassage(input: { reference: string, html?: boolean }) {
  const rawReference = String(input?.reference || '').trim();
  const reference = normalizeReferenceSpacing(rawReference);
  if (!reference) return { reference, text: '' };

  const props = PropertiesService.getScriptProperties();
  const token = String(props.getProperty('ESV_API_TOKEN') || '');
  if (!token) {
    return { reference, text: '', html: '', error: 'ESV_API_TOKEN not set in Script Properties' };
  }

  const includeHtml = input?.html !== false;
  const splitRefs = splitReferenceIntoDistinctBooks(rawReference);
  const multiRefs = splitRefs.length ? splitRefs : [];
  const refsToFetch = multiRefs.length ? multiRefs : [reference];
  const includeInlineReference = !multiRefs.length;

  const chunks = refsToFetch.map(ref => fetchPassageChunk(ref, token, includeHtml, includeInlineReference));

  if (!multiRefs.length) {
    const first = chunks[0] || { reference, text: '', html: '' };
    return { reference, text: first.text, html: first.html };
  }

  const textBlocks = chunks.map(chunk => formatChunkText(chunk.reference, chunk.text)).filter(Boolean);
  const text = textBlocks.join('\n\n').trim();
  let html = '';
  if (includeHtml) {
    html = chunks.map(chunk => formatChunkHtml(chunk.reference, chunk.html, chunk.text)).filter(Boolean).join('');
  }
  return { reference, text, html };
}

export function lblaPassage(input: { reference: string }): PassageResult {
  const rawReference = String(input?.reference || '').trim();
  const reference = normalizeReferenceSpacing(rawReference);
  if (!reference) return { reference, text: '' };

  try {
    const splitRefs = splitReferenceIntoDistinctBooks(rawReference);
    const refsToFetch = splitRefs.length ? splitRefs : [reference];
    const chunks = refsToFetch.map(ref => fetchBibleGatewayChunk(ref, 'LBLA'));
    if (!splitRefs.length) {
      const first = chunks[0] || { reference, text: '' };
      return { reference, text: first.text };
    }
    const text = chunks.map(chunk => formatChunkText(chunk.reference, chunk.text)).filter(Boolean).join('\n\n').trim();
    return { reference, text };
  } catch (err) {
    const message = err && (err as any).message ? String((err as any).message) : 'Unable to fetch LBLA passage';
    return { reference, text: '', error: message };
  }
}

export function getScriptureVersions(input: { reference: string }) {
  const esv = esvPassage({ reference: String(input?.reference || ''), html: false });
  const lbla = lblaPassage({ reference: String(input?.reference || '') });
  return {
    reference: esv.reference || lbla.reference || normalizeReferenceSpacing(String(input?.reference || '')),
    esv,
    lbla
  };
}
