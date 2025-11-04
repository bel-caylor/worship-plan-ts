// src/features/songs.ts
import {
  SONG_SHEET, SONG_COL_NAME, FOLDER_LINK_COL, AUDIO_LINKS_COL, MAX_AUDIO_LINKS,
  ROOT_FOLDER_ID, SPANISH_ROOT_ID, SP_COL_NAME, TARGET_LEADER_COL, Row,
  PLANNER_SHEET, PLANNER_SONG_COLS
} from '../constants';
import { getSheetByName, getHeaders, ensureColumn } from '../util/sheets';
import { findBestFolderForSong, listAudioInFolder } from '../util/drive';
import { splitTokens } from '../util/text';


export function linkSongMedia() {
    const sh = getSheetByName(SONG_SHEET);
    const { headers, colMap } = getHeaders(sh);
    ensureColumn(sh, headers, colMap, FOLDER_LINK_COL);
    ensureColumn(sh, headers, colMap, AUDIO_LINKS_COL);

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const rng = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn());
    const values = rng.getValues();

    // Prepare a RichText builder range for audio links
    const audioCol = colMap[AUDIO_LINKS_COL] + 1; // 1-based col index in sheet
    const folderCol = colMap[FOLDER_LINK_COL] + 1;

    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        for (let i = 0; i < values.length; i++) {
            const row = values[i];
            const song = String(row[colMap[SONG_COL_NAME]] || '').trim();
            if (!song) continue;

            // Find best matching folder for this song
            const isSpanish = String(row[colMap[SP_COL_NAME]] || '').trim().toUpperCase() === 'Y';
            const rootIdForRow = isSpanish ? SPANISH_ROOT_ID : ROOT_FOLDER_ID;
            const match = findBestFolderForSong(song, rootIdForRow);
            let folderUrl = '';
            let audioRich: GoogleAppsScript.Spreadsheet.RichTextValue | null = null;

            if (match) {
                folderUrl = match.url;

                // List audio files in that folder
                const audioFiles = listAudioInFolder(match.id, MAX_AUDIO_LINKS);
                if (audioFiles.length) {
                    const builder = SpreadsheetApp.newRichTextValue();
                    let text = '';
                    const runs: { start: number; end: number; url: string }[] = [];
                    audioFiles.forEach((f, idx) => {
                        const label = f.name;
                        const start = text.length;
                        text += label + (idx < audioFiles.length - 1 ? '\n' : '');
                        runs.push({ start, end: start + label.length, url: f.url });
                    });
                    builder.setText(text);
                    runs.forEach(r => builder.setLinkUrl(r.start, r.end, r.url));
                    audioRich = builder.build();
                }
            }

            // Write to the sheet (Folder URL as hyperlink; Audio as rich text)
            const rowIndex = 2 + i;

            if (folderUrl) {
                sh.getRange(rowIndex, folderCol).setFormula(`=HYPERLINK("${folderUrl}","Open Folder")`);
            } else {
                sh.getRange(rowIndex, folderCol).clearContent();
            }

            const audioCell = sh.getRange(rowIndex, audioCol);
            if (audioRich) {
                audioCell.setRichTextValue(audioRich);
                audioCell.setWrap(true);
            } else {
                audioCell.clearContent();
            }
        }
    } finally {
        lock.releaseLock();
    }

    SpreadsheetApp.getActive().toast('Linking complete', 'Worship Planner', 4);
}

export function getSongsWithLinksForView(): Row[] {
    const sh = getSheetByName(SONG_SHEET);

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return [];

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
        .getValues()[0]
        .map(h => String(h ?? '').trim());

    const colMap: Record<string, number> = {};
    headers.forEach((h, i) => (colMap[h] = i));

    const fCol = colMap[FOLDER_LINK_COL] ?? -1;
    const leaderColIdx =
        (colMap[TARGET_LEADER_COL] ?? -1) >= 0
            ? (colMap[TARGET_LEADER_COL] as number)
            : (colMap['Leader'] ?? -1);

    const bodyRange = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn());
    const values = bodyRange.getValues();
    const formulas = bodyRange.getFormulas();
    const rich = bodyRange.getRichTextValues();

    const out: Row[] = [];

    for (let r = 0; r < values.length; r++) {
        const rowObj: Row = {};
        headers.forEach((h, c) => (rowObj[h] = values[r][c]));

        // ---- Folder URL ----
        let folderUrl: string | null = null;
        if (fCol >= 0) {
            const rt = rich[r][fCol];
            try {
                // @ts-ignore
                const url = rt && typeof rt.getLinkUrl === 'function' ? rt.getLinkUrl() : null;
                if (url) folderUrl = url;
            } catch { }
            if (!folderUrl) {
                const f = formulas[r][fCol];
                const m = /^=HYPERLINK\("([^"]+)"/i.exec(String(f || ''));
                if (m) folderUrl = m[1];
            }
        }
        (rowObj as any)._folderUrl = folderUrl || null;

        // ---- Leaders array for filtering ----
        const leadersRaw = leaderColIdx >= 0 ? String(values[r][leaderColIdx] ?? '') : '';
        (rowObj as any)._leaders = splitTokens(leadersRaw);

        (rowObj as any)._audio = [];
        out.push(rowObj);
    }

    return out;
}

// Return selected fields for specific song names (robust matching)
export function getSongFields(input: { names?: string[]; fields?: string[] }) {
  const names = Array.isArray(input?.names) ? input!.names.map(String) : [];
  const want = Array.isArray(input?.fields) && input!.fields!.length ? input!.fields!.map(String) : ['Lyrics','Themes'];
  if (!names.length) return { items: [] };

  const sh = getSheetByName(SONG_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { items: [] };
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v ?? '').trim());
  const H = headers.map(h => h.toLowerCase());
  const idx = (label: string) => H.indexOf(label.toLowerCase());
  const nameIdx = idx('song');
  const colIdx: Record<string, number> = {};
  for (const f of want) {
    const i = idx(f);
    if (i >= 0) colIdx[f] = i;
  }
  const body = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  const norm = (s: string) => String(s || '')
    .toLowerCase()
    .replace(/\([^)]*\)|\[[^\]]*\]/g, ' ')
    .replace(/\+sp|\+es/g, ' ')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const rows = body.map(r => ({
    title: nameIdx >= 0 ? String(r[nameIdx] ?? '') : '',
    normTitle: nameIdx >= 0 ? norm(String(r[nameIdx] ?? '')) : '',
    r
  }));

  const out: { name: string; fields: Record<string,string> }[] = [];
  for (const raw of names) {
    const q = norm(String(raw || ''));
    let best: any = null; let bestScore = -1;
    for (const row of rows) {
      if (!row.normTitle) continue;
      let score = 0;
      if (row.normTitle === q) score = 2;
      else if (row.normTitle.includes(q) || q.includes(row.normTitle)) {
        score = Math.min(row.normTitle.length, q.length) / Math.max(row.normTitle.length, q.length);
      } else {
        const A = new Set(row.normTitle.split(' '));
        const B = new Set(q.split(' '));
        let inter = 0; for (const x of A) if (B.has(x)) inter++;
        const union = A.size + B.size - inter;
        score = union ? inter / union : 0;
      }
      if (score > bestScore) { best = row; bestScore = score; }
    }
    const fields: Record<string,string> = {};
    if (best) {
      for (const f of Object.keys(colIdx)) {
        const v = best.r[colIdx[f]];
        fields[f] = v != null ? String(v) : '';
      }
    }
    out.push({ name: String(raw || ''), fields });
  }
  return { items: out };
}

// --- Suggestions ---
type SuggestInput = {
  theme?: string;
  scripture?: string; // raw reference or text
  slot?: string;      // e.g., 'Opening Song','Song 1','Communion Song'
  k?: number;
};

export function suggestSongs(input: SuggestInput) {
  const k = Math.max(1, Math.min(20, Number(input?.k ?? 5)));
  const slot = String(input?.slot || '').toLowerCase();
  const theme = String(input?.theme || '').toLowerCase();
  const scripture = String(input?.scripture || '').toLowerCase();
  const queryTokens = tokenSet(theme + ' ' + scripture);

  // Slot→keywords mapping
  const slotTags = new Set<string>((() => {
    if (slot.includes('communion')) return ['communion','lord supper','eucharist'];
    if (slot.includes('offering')) return ['offering','giving','stewardship'];
    if (slot.includes('opening') || /\bsong\s*1\b/.test(slot) || slot.includes('call to worship')) return ['call to worship','praise','gather','welcome'];
    if (slot.includes('closing') || slot.includes('benediction')) return ['benediction','sending','go','dismissal'];
    return [];
  })());

  const sh = getSheetByName(SONG_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { items: [] };
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v ?? ''));
  const col = (cands: string[]) => {
    const H = headers.map(h => h.trim().toLowerCase());
    for (const c of cands) {
      const i = H.indexOf(c.toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  };
  const nameIdx = col(['song','title','name', 'song title']);
  const keywordsIdx = col(['keywords','tags']);
  const themesIdx = col(['themes','theme']);
  const seasonsIdx = col(['seasons','season']);
  const scripturesIdx = col(['scriptures','scripture refs','scripture']);
  const lyricsIdx = col(['lyrics','text']);
  const lastUsedIdx = col(['last used','lastused','last_used']);
  const usageIdx = col(['usage','uses','count']);

  const body = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  const scored: { name: string; score: number; reason: string }[] = [];
  for (const row of body) {
    const name = nameIdx >= 0 ? String(row[nameIdx] ?? '') : '';
    if (!name) continue;
    const kw = keywordsIdx >= 0 ? tokenSet(String(row[keywordsIdx] ?? '')) : new Set<string>();
    const th = themesIdx >= 0 ? tokenSet(String(row[themesIdx] ?? '')) : new Set<string>();
    const ssn = seasonsIdx >= 0 ? tokenSet(String(row[seasonsIdx] ?? '')) : new Set<string>();
    const scr = scripturesIdx >= 0 ? tokenSet(String(row[scripturesIdx] ?? '')) : new Set<string>();
    const lyr = lyricsIdx >= 0 ? lightTokenSet(String(row[lyricsIdx] ?? '')) : new Set<string>();
    const lastUsed = lastUsedIdx >= 0 ? row[lastUsedIdx] : '';
    const usage = usageIdx >= 0 ? Number(row[usageIdx] ?? 0) : 0;

    let score = 0;
    const hits: string[] = [];

    // Slot tag boost
    for (const t of slotTags) if (hasToken(kw, t) || hasToken(th, t)) { score += 3; hits.push(`slot:${t}`); }

    // Keyword/Theme/Season/Scripture overlaps
    const overlap = overlapCount(kw, queryTokens) * 2 + overlapCount(th, queryTokens) * 1.5 + overlapCount(ssn, queryTokens) * 1 + overlapCount(scr, queryTokens) * 1.5;
    if (overlap) { score += overlap; hits.push('query'); }

    // Lyrics light hit
    const lyrOverlap = overlapCount(lyr, queryTokens);
    if (lyrOverlap) { score += Math.min(2, lyrOverlap * 0.5); hits.push('lyrics'); }

    // Recency penalty
    let rec = 0;
    try {
      if (lastUsed instanceof Date && !isNaN(lastUsed.getTime())) {
        const days = Math.floor((Date.now() - lastUsed.getTime()) / (1000*60*60*24));
        if (days < 28) rec = -2; else if (days < 56) rec = -1;
      }
    } catch {}
    score += rec;

    // Heavy usage penalty
    if (usage > 0) score += Math.max(-2, -usage * 0.1);

    if (score > 0) scored.push({ name, score, reason: hits.join(',') || 'match' });
  }

  scored.sort((a,b) => b.score - a.score);
  return { items: scored.slice(0, k) };
}

function tokenSet(s: string) {
  return new Set(String(s || '').toLowerCase().split(/[^a-z0-9\+]+/).filter(Boolean));
}
function lightTokenSet(s: string) {
  return new Set(String(s || '').toLowerCase().split(/[^a-z0-9]+/).filter(w => w.length >= 5));
}
function hasToken(set: Set<string>, token: string) {
  const t = String(token || '').toLowerCase();
  if (!t) return false;
  if (set.has(t)) return true;
  // also allow exact phrase in a joined string
  return false;
}
function overlapCount(setA: Set<string>, setB: Set<string>) {
  let c = 0; for (const t of setB) if (setA.has(t)) c++; return c;
}


export function rebuildSongUsageFromPlanner() {
  const planner = getSheetByName(PLANNER_SHEET);
  const pLastRow = planner.getLastRow();
  const pLastCol = planner.getLastColumn();
  if (pLastRow < 2 || pLastCol < 1) return { updated: 0 };

  const pHeaders = planner.getRange(1, 1, 1, pLastCol).getValues()[0].map(v => String(v ?? '').trim());
  const pIdx = (name: string) => pHeaders.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const pColMap: Record<string, number> = {};
  for (const label of PLANNER_SONG_COLS) {
    const i = pIdx(label);
    if (i >= 0) pColMap[label] = i;
  }
  if (!Object.keys(pColMap).length) return { updated: 0 };

  const pBody = planner.getRange(2, 1, pLastRow - 1, pLastCol).getValues();
  const norm = (s: string) => String(s || '').trim().replace(/\s+/g, ' ');
  const songUsage = new Map<string, Set<string>>();
  const labelMap: Record<string, string> = {
    'Opening Song': 'Call to Worship',
    'Song2': 'Song2',
    'Song3': 'Song3',
    'Song4/Communion': 'Song4',
    'Offering/Communion Song': 'Communion',
    'Closing Song': 'Closing'
  };

  for (const row of pBody) {
    for (const label of PLANNER_SONG_COLS) {
      const idx = pColMap[label];
      if (idx == null || idx < 0) continue;
      const raw = norm(String(row[idx] ?? ''));
      if (!raw) continue;
      const key = norm(raw);
      if (!songUsage.has(key)) songUsage.set(key, new Set());
      songUsage.get(key)!.add(labelMap[label] || label);
    }
  }

  const songsSh = getSheetByName(SONG_SHEET);
  const { headers, colMap } = getHeaders(songsSh);
  ensureColumn(songsSh, headers, colMap, 'Usage');
  const lastRow = songsSh.getLastRow();
  if (lastRow < 2) return { updated: 0 };
  const lastCol = songsSh.getLastColumn();
  const nameCol = colMap[SONG_COL_NAME];
  const usageCol = colMap['Usage'];

  const range = songsSh.getRange(2, 1, lastRow - 1, lastCol);
  const values = range.getValues();
  let updated = 0;
  const order = ['Call to Worship','Song2','Song3','Song4','Communion','Closing'];
  for (let r = 0; r < values.length; r++) {
    const name = norm(String(values[r][nameCol] ?? ''));
    let usage = '';
    if (name && songUsage.has(name)) {
      const arr = Array.from(songUsage.get(name)!.values());
      arr.sort((a, b) => order.indexOf(a) - order.indexOf(b));
      usage = arr.join(', ');
    }
    if (String(values[r][usageCol] ?? '') !== usage) {
      values[r][usageCol] = usage;
      updated++;
    }
  }
  range.setValues(values);
  return { updated };
}
