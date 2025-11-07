// src/features/songs.ts
import {
  SONG_SHEET, SONG_COL_NAME, FOLDER_LINK_COL, AUDIO_LINKS_COL, MAX_AUDIO_LINKS,
  ROOT_FOLDER_ID, SPANISH_ROOT_ID, SP_COL_NAME, TARGET_LEADER_COL, Row,
  PLANNER_SHEET, PLANNER_SONG_COLS
} from '../constants';
import { getSheetByName, getHeaders, ensureColumn } from '../util/sheets';
import { findBestFolderForSong, listAudioInFolder } from '../util/drive';
import { splitTokens } from '../util/text';
import { aiSongMetadata } from '../util/ai';

type UpdateSongUsageInput = {
  name?: string;
  date?: string;
  incrementUses?: boolean;
};

function normalizeSongTitle(s: string) {
  return String(s || '')
    .toLowerCase()
    .replace(/\([^)]*\)|\[[^\]]*\]/g, ' ')
    .replace(/\+sp|\+es/g, ' ')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function toSheetDate(input: string) {
  const s = String(input || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const [yy, mm, dd] = s.split('-').map(Number);
    return new Date(Date.UTC(yy, mm - 1, dd));
  }
  try {
    const d = new Date(s);
    if (!isNaN(d.getTime())) {
      return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    }
  } catch (_) { /* ignore */ }
  return s;
}

export function updateSongRecency(input: UpdateSongUsageInput) {
  const nameRaw = String(input?.name || '').trim();
  if (!nameRaw) throw new Error('Song name required');

  const sh = getSheetByName(SONG_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { updated: false };

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (label: string) => headers.findIndex(h => h.toLowerCase() === label.toLowerCase());

  const nameIdx = col(SONG_COL_NAME) >= 0 ? col(SONG_COL_NAME) : headers.findIndex(h => /song/i.test(h));
  if (nameIdx < 0) throw new Error('Song column not found');

  const lastUsedIdx = (() => {
    const direct = col('Last_Used');
    if (direct >= 0) return direct;
    const spaced = col('Last Used');
    if (spaced >= 0) return spaced;
    return headers.findIndex(h => h.toLowerCase().includes('last') && h.toLowerCase().includes('used'));
  })();
  const usesIdx = (() => {
    const plain = col('Uses');
    if (plain >= 0) return plain;
    return headers.findIndex(h => h.toLowerCase() === 'use count' || h.toLowerCase().includes('uses'));
  })();
  const yearsIdx = (() => {
    const plain = col('Years_Used');
    if (plain >= 0) return plain;
    const spaced = col('Years Used');
    if (spaced >= 0) return spaced;
    return headers.findIndex(h => h.toLowerCase().includes('years') && h.toLowerCase().includes('used'));
  })();

  const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
  const values = dataRange.getValues();

  const target = normalizeSongTitle(nameRaw);
  let matchIdx = -1;
  for (let r = 0; r < values.length; r++) {
    const rawName = String(values[r][nameIdx] ?? '').trim();
    if (!rawName) continue;
    if (normalizeSongTitle(rawName) === target) {
      matchIdx = r;
      break;
    }
  }

  if (matchIdx < 0) return { updated: false };

  const absoluteRow = matchIdx + 2;
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    let lastUsedDate = '';
    if (lastUsedIdx >= 0 && input?.date) {
      const val = toSheetDate(input.date);
      sh.getRange(absoluteRow, lastUsedIdx + 1).setValue(val);
      lastUsedDate = String(input.date || '');
    }
    let usesCount: number | undefined;
    if (usesIdx >= 0 && (input?.incrementUses ?? true)) {
      const current = Number(values[matchIdx][usesIdx]) || 0;
      usesCount = current + 1;
      sh.getRange(absoluteRow, usesIdx + 1).setValue(usesCount);
    }
    if (yearsIdx >= 0 && input?.date) {
      const current = String(values[matchIdx][yearsIdx] ?? '').trim();
      const year = String(input.date).slice(0, 4);
      if (year && /^\d{4}$/.test(year)) {
        const parts = current ? current.split(',').map(s => s.trim()).filter(Boolean) : [];
        if (!parts.includes(year)) {
          parts.push(year);
          parts.sort((a, b) => Number(a) - Number(b));
          sh.getRange(absoluteRow, yearsIdx + 1).setValue(parts.join(', '));
        }
      }
    }
    return {
      updated: true,
      lastUsed: lastUsedDate || input?.date || '',
      uses: usesCount
    };
  } finally {
    lock.releaseLock();
  }

  return { updated: true, lastUsed: input?.date || '' };
}


export function linkSongMedia() {
    const sh = getSheetByName(SONG_SHEET);
    const details = getHeaders(sh);
    const headers = details.headers;
    const colMap = details.colMap;

    if (!Object.prototype.hasOwnProperty.call(colMap, FOLDER_LINK_COL)) {
        ensureColumn(sh, headers, colMap, FOLDER_LINK_COL);
    }

    const folderColIdx = colMap[FOLDER_LINK_COL] ?? -1;
    const audioColIdx = colMap[AUDIO_LINKS_COL] ?? -1;
    const songColIdx = colMap[SONG_COL_NAME] ?? -1;
    const spColIdx = colMap[SP_COL_NAME] ?? -1;

    if (songColIdx < 0) {
        SpreadsheetApp.getActive().toast('Songs sheet is missing Song column.', 'Worship Planner', 4);
    return;
}


    if (folderColIdx < 0 && audioColIdx < 0) {
        SpreadsheetApp.getActive().toast('No media columns present; nothing to update.', 'Worship Planner', 4);
        return;
    }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const rng = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn());
    const values = rng.getValues();

    const folderCol = folderColIdx >= 0 ? folderColIdx + 1 : null;
    const audioCol = audioColIdx >= 0 ? audioColIdx + 1 : null;

    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
        for (let i = 0; i < values.length; i++) {
            const row = values[i];
            const song = String(row[songColIdx] || '').trim();
            if (!song) continue;

            const isSpanish = spColIdx >= 0
                ? String(row[spColIdx] || '').trim().toUpperCase() === 'Y'
                : false;
            const rootIdForRow = isSpanish ? SPANISH_ROOT_ID : ROOT_FOLDER_ID;
            const match = findBestFolderForSong(song, rootIdForRow);
            let folderUrl = '';
            let audioRich: GoogleAppsScript.Spreadsheet.RichTextValue | null = null;

            if (match) {
                folderUrl = match.url;

                if (audioCol != null) {
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
            }

            const rowIndex = 2 + i;

            if (folderCol != null) {
                if (folderUrl) {
                    sh.getRange(rowIndex, folderCol).setFormula(`=HYPERLINK("${folderUrl}","Open Folder")`);
                } else {
                    sh.getRange(rowIndex, folderCol).clearContent();
                }
            }

            if (audioCol != null) {
                const audioCell = sh.getRange(rowIndex, audioCol);
                if (audioRich) {
                    audioCell.setRichTextValue(audioRich);
                    audioCell.setWrap(true);
                } else {
                    audioCell.clearContent();
                }
            }
        }
    } finally {
        lock.releaseLock();
    }

    SpreadsheetApp.getActive().toast('Linking complete', 'Worship Planner', 4);
}


type FolderFileDetails = {
  id: string;
  name: string;
  url: string;
  mimeType: string;
  created: Date | null;
};

type FolderCandidate = {
  id: string;
  name: string;
  url: string;
  path: string[];
  files: FolderFileDetails[];
  flags: {
    isSpanish: boolean;
    forcedSeason?: string;
    isChristmas?: boolean;
    isArchive?: boolean;
    inAdventCollection?: boolean;
  };
};

type SyncOptions = {
  dryRun?: boolean;
  limit?: number;
  reset?: boolean;
  maxRuntimeMs?: number;
};

export function syncSongsFromDrive(options?: SyncOptions) {
  const sh = getSheetByName(SONG_SHEET);
  const details = getHeaders(sh);
  const headers = details.headers;
  const colMap = details.colMap;
  const required = [FOLDER_LINK_COL, 'Lyrics', 'Themes', 'Season', 'Keywords', 'Scriptures', 'First_Used', SP_COL_NAME, 'Archive'];
  for (const label of required) ensureColumn(sh, headers, colMap, label);
  const lastCol = headers.length;
  const lastRow = sh.getLastRow();
  let existingValues: any[][] = [];
  let existingFormulas: string[][] = [];
  if (lastRow > 1) {
    const existingRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    existingValues = existingRange.getValues();
    existingFormulas = existingRange.getFormulas();
  }
  const existingIndex = buildExistingSongIndex(existingValues, existingFormulas, colMap);
  const existingFolderIds = buildExistingFolderIdSet(existingValues, existingFormulas, colMap);

  const candidates = gatherSongFolderCandidates();
  const limit = Math.max(0, Number(options?.limit || 0));
  const pending: any[][] = [];
  let added = 0;
  let skipped = 0;
  let processed = 0;
  const errors: string[] = [];
  const stateKey = 'SONG_SYNC_CURSOR';
  const scriptProps = PropertiesService.getScriptProperties();
  const previousState = options?.reset ? null : loadSyncState(scriptProps, stateKey);
  const resumeId = previousState?.lastId || '';
  const startIndex = resumeId ? Math.max(0, candidates.findIndex(c => c.id === resumeId) + 1) : 0;
  const maxRuntime = Math.max(60_000, Number(options?.maxRuntimeMs || (4.5 * 60 * 1000)));
  const startedAt = Date.now();

  const total = candidates.length;
  const stepRangeStart = Math.max(2, lastRow + 1);
  let nextWriteRow = stepRangeStart;
  const flushRows = () => {
    if (!pending.length || options?.dryRun) return;
    const rng = sh.getRange(nextWriteRow, 1, pending.length, lastCol);
    rng.setValues(pending.splice(0, pending.length));
    nextWriteRow += rng.getNumRows();
  };

  let resumeNeeded = false;
  let lastProcessedId = resumeId;
  for (let idx = startIndex; idx < candidates.length; idx++) {
    const candidate = candidates[idx];
    const elapsed = Date.now() - startedAt;
    if (elapsed >= maxRuntime) {
      resumeNeeded = true;
      break;
    }
    if (limit && added >= limit) break;
    processed++;

    const displayName = cleanSongTitle(candidate.name);
    if (!displayName) continue;

    const normName = normalizeSongTitle(displayName);
    const folderId = candidate.id;
    const existing = existingIndex.get(normName);
    if ((existing && existing.has(folderId)) || existingFolderIds.has(folderId)) {
      skipped++;
      lastProcessedId = candidate.id;
      continue;
    }

    const bestFile = pickLyricsFile(candidate.files);
    const lyrics = bestFile ? readLyricsFromFile(bestFile) : '';
    const hints: string[] = [];
    if (candidate.path.length) hints.push(`Folder path: ${candidate.path.join(' / ')}`);
    if (candidate.flags?.isSpanish) hints.push('Spanish language');
    if (candidate.flags?.isChristmas) hints.push('Christmas / Advent song');
    if (bestFile) hints.push(`Source file: ${bestFile.name}`);
    const forcedSeason = candidate.flags?.forcedSeason || (candidate.flags?.isChristmas ? 'Christmas' : '');
    const metadata = aiSongMetadata({
      name: displayName,
      lyrics,
      hints,
      forcedSeason: forcedSeason || undefined,
      kScriptures: 5,
      allowKeywords: !!lyrics
    });
    if (metadata.error) errors.push(`${displayName}: ${metadata.error}`);

    const firstUsed = oldestFileDate(candidate.files);
    const row = new Array(lastCol).fill('');
    row[colMap[SONG_COL_NAME]] = displayName;
    if (colMap[FOLDER_LINK_COL] != null) row[colMap[FOLDER_LINK_COL]] = makeFolderHyperlink(candidate.url);
    if (colMap['Lyrics'] != null) row[colMap['Lyrics']] = lyrics;
    if (colMap['Themes'] != null) row[colMap['Themes']] = metadata.themes.join(', ');
    const seasonRaw = forcedSeason || metadata.season || '';
    const seasonValue = (candidate.flags?.inAdventCollection || /christmas/i.test(seasonRaw)) ? 'Christmas' : 'General';
    if (colMap['Season'] != null) row[colMap['Season']] = seasonValue;
    if (colMap['Archive'] != null && candidate.flags?.isArchive) row[colMap['Archive']] = 'Y';
    if (colMap['Keywords'] != null) row[colMap['Keywords']] = metadata.keywords.join(', ');
    if (colMap['Scriptures'] != null) row[colMap['Scriptures']] = metadata.scriptures.join(', ');
    if (colMap['First_Used'] != null) row[colMap['First_Used']] = firstUsed || '';
    if (colMap[SP_COL_NAME] != null && candidate.flags?.isSpanish) row[colMap[SP_COL_NAME]] = 'Y';

    pending.push(row);
    if (pending.length >= 20) flushRows();
    added++;
    if (!existingIndex.has(normName)) existingIndex.set(normName, new Set());
    existingIndex.get(normName)!.add(folderId);
    existingFolderIds.add(folderId);
    lastProcessedId = candidate.id;
  }

  flushRows();

  if (resumeNeeded && processed === 0) {
    // If we resumed but processed nothing (e.g., startIndex beyond list), reset.
    resumeNeeded = false;
    lastProcessedId = '';
  }
  if (resumeNeeded) {
    saveSyncState(scriptProps, stateKey, { lastId: lastProcessedId || candidates[candidates.length - 1]?.id || '', at: Date.now() });
  } else {
    clearSyncState(scriptProps, stateKey);
  }

  const summary = {
    added,
    skipped,
    processed,
    scanned: total,
    dryRun: !!options?.dryRun,
    resumeNeeded,
    resumeId: resumeNeeded ? lastProcessedId : undefined,
    errors
  };
  try {
    const suffix = resumeNeeded ? ' (paused, run again to continue)' : '';
    SpreadsheetApp.getActive().toast(`Songs sync: +${added} / skipped ${skipped}${suffix}`);
  } catch (_) {}
  return summary;
}

function loadSyncState(props: GoogleAppsScript.Properties.Properties, key: string) {
  try {
    const raw = props.getProperty(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (_) {
    return null;
  }
}

function saveSyncState(props: GoogleAppsScript.Properties.Properties, key: string, state: Record<string, unknown>) {
  try { props.setProperty(key, JSON.stringify(state)); } catch (_) {}
}

function clearSyncState(props: GoogleAppsScript.Properties.Properties, key: string) {
  try { props.deleteProperty(key); } catch (_) {}
}

function buildExistingSongIndex(rows: any[][], formulas: string[][], colMap: Record<string, number>) {
  const out = new Map<string, Set<string>>();
  const nameIdx = colMap[SONG_COL_NAME];
  const folderIdx = colMap[FOLDER_LINK_COL];
  if (nameIdx == null) return out;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const title = String(row[nameIdx] ?? '').trim();
    if (!title) continue;
    const norm = normalizeSongTitle(title);
    if (!out.has(norm)) out.set(norm, new Set());
    const formulaRow = Array.isArray(formulas?.[r]) ? formulas[r] : [];
    const folderFormula = folderIdx != null ? String(formulaRow[folderIdx] ?? '') : '';
    const folderValue = folderIdx != null ? row[folderIdx] : '';
    const folderUrl = extractFolderUrl(folderValue, folderFormula);
    const folderId = folderUrl ? extractFolderId(folderUrl) : '';
    out.get(norm)!.add(folderId || '');
  }
  return out;
}

function buildExistingFolderIdSet(rows: any[][], formulas: string[][], colMap: Record<string, number>) {
  const out = new Set<string>();
  const folderIdx = colMap[FOLDER_LINK_COL];
  if (folderIdx == null) return out;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const formulaRow = Array.isArray(formulas?.[r]) ? formulas[r] : [];
    const folderFormula = String(formulaRow[folderIdx] ?? '');
    const folderValue = row[folderIdx];
    const folderUrl = extractFolderUrl(folderValue, folderFormula);
    const folderId = folderUrl ? extractFolderId(folderUrl) : '';
    if (folderId) out.add(folderId);
  }
  return out;
}

type FolderStackItem = {
  folder: GoogleAppsScript.Drive.Folder;
  path: string[];
  isSpanish: boolean;
  forcedSeason?: string;
  depth: number;
};

const ALLOWED_PLUS_FOLDERS = new Set(['+spanish', '+archive', '+advent/christmas']);

function gatherSongFolderCandidates(): FolderCandidate[] {
  const roots = [
    { id: ROOT_FOLDER_ID, isSpanish: false },
    { id: SPANISH_ROOT_ID, isSpanish: true }
  ];
  const seen = new Set<string>();
  const out: FolderCandidate[] = [];
  for (const root of roots) {
    if (!root.id) continue;
    let folder: GoogleAppsScript.Drive.Folder;
    try {
      folder = DriveApp.getFolderById(root.id);
    } catch (_) {
      continue;
    }
    const stack: FolderStackItem[] = [{
      folder,
      path: [folder.getName()],
      isSpanish: root.isSpanish,
      forcedSeason: undefined,
      depth: 0
    }];
    while (stack.length) {
      const current = stack.pop()!;
      const name = current.folder.getName();
      const normalizedPath = current.path.map(p => p.toLowerCase());
      const isChristmas = normalizedPath.some(seg => seg.includes('advent') || seg.includes('christmas') || seg.includes('navidad'));
      const isSpanish = current.isSpanish || normalizedPath.some(seg => seg.includes('spanish') || seg.includes('espanol') || seg.includes('español'));
      const inAdventCollection = normalizedPath.some(seg => seg === '+advent/christmas');
      const forcedSeason = current.forcedSeason || (inAdventCollection || isChristmas ? 'Christmas' : undefined);
      const isArchive = normalizedPath.some(seg => seg === '+archive');
      const files = listFolderFiles(current.folder);
      const include = current.depth > 0 && !isCategoryFolder(name) && files.length > 0;
      if (include) {
        const id = current.folder.getId();
        if (!seen.has(id)) {
          seen.add(id);
          out.push({
            id,
            name,
            url: current.folder.getUrl(),
            path: current.path.slice(1),
            files,
            flags: {
              isSpanish,
              forcedSeason,
              isChristmas,
              isArchive,
              inAdventCollection
            }
          });
        }
      }
      const children = listChildFolders(current.folder);
      for (const child of children) {
        stack.push({
          folder: child,
          path: [...current.path, child.getName()],
          isSpanish,
          forcedSeason,
          depth: current.depth + 1
        });
      }
    }
  }
  out.sort((a, b) => a.name.localeCompare(b.name));
  return out;
}

function listChildFolders(folder: GoogleAppsScript.Drive.Folder) {
  const out: GoogleAppsScript.Drive.Folder[] = [];
  const it = folder.getFolders();
  while (it.hasNext()) out.push(it.next());
  return out;
}

function listFolderFiles(folder: GoogleAppsScript.Drive.Folder) {
  const out: FolderFileDetails[] = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const file = it.next();
    out.push({
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      mimeType: file.getMimeType(),
      created: file.getDateCreated()
    });
  }
  return out;
}

function isCategoryFolder(name: string) {
  const trimmed = String(name || '').trim();
  if (!trimmed) return true;
  const lower = trimmed.toLowerCase();
  if (trimmed.startsWith('+') && !ALLOWED_PLUS_FOLDERS.has(lower)) return true;
  return false;
}

function cleanSongTitle(raw: string) {
  let name = String(raw || '').replace(/\s+/g, ' ').trim();
  name = name.replace(/^\d+[\s._-]+/, '').trim();
  if (!name) return String(raw || '').trim();
  return name;
}

function pickLyricsFile(files: FolderFileDetails[]) {
  let best: FolderFileDetails | null = null;
  let bestScore = -1;
  for (const file of files) {
    const score = scoreLyricsFile(file);
    if (score > bestScore) {
      best = file;
      bestScore = score;
    }
  }
  return best;
}

function scoreLyricsFile(file: FolderFileDetails) {
  const name = String(file?.name || '').toLowerCase();
  const mime = String(file?.mimeType || '').toLowerCase();
  let score = 0;
  if (!name) return 0;
  if (mime === 'application/vnd.google-apps.document') score += 8;
  if (/_w/i.test(name) || /-w/i.test(name)) score += 5;
  if (/lyric/i.test(name) || /letras/i.test(name)) score += 4;
  if (/(^|[\s_-])w($|[\s_-])/i.test(file.name)) score += 2;
  if (/words?/.test(name) || /palabras/.test(name)) score += 1;
  if (mime.startsWith('text/')) score += 1;
  return score;
}

function readLyricsFromFile(file: FolderFileDetails) {
  if (!file) return '';
  const mime = String(file.mimeType || '').toLowerCase();
  try {
    if (mime === 'application/vnd.google-apps.document') {
      const doc = DocumentApp.openById(file.id);
      const text = doc.getBody().getText();
      return cleanLyricsText(text, file.name);
    }
    if (mime.startsWith('text/')) {
      const blob = DriveApp.getFileById(file.id).getBlob();
      return cleanLyricsText(blob.getDataAsString(), file.name);
    }
  } catch (err) {
    try { Logger.log(`Lyrics read error (${file.name}): ${(err as any)?.message}`); } catch (_) {}
  }
  return '';
}

function cleanLyricsText(text: string, songName: string) {
  const titleNorm = normalizeSongTitle(songName);
  const lines = String(text || '')
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map(l => l.replace(/\uFEFF/g, '').trim());
  const out: string[] = [];
  for (const line of lines) {
    if (!line && !out.length) continue;
    const norm = normalizeSongTitle(line);
    if (!out.length) {
      if (!line) continue;
      if (norm && norm === titleNorm) continue;
      if (/^title[:\s]/i.test(line)) continue;
      if (/^lyrics?$/i.test(line)) continue;
      if (/^words? and music/i.test(line)) continue;
    }
    if (/ccli|copyright|all rights reserved/i.test(line)) continue;
    out.push(line);
  }
  const clean = out.join('\n').replace(/\n{3,}/g, '\n\n').trim();
  return clean;
}

function oldestFileDate(files: FolderFileDetails[]) {
  let best: Date | null = null;
  for (const file of files) {
    if (!file.created) continue;
    if (!best || file.created.getTime() < best.getTime()) {
      best = file.created;
    }
  }
  return best;
}

function extractFolderId(url: string) {
  if (!url) return '';
  const direct = /\/folders\/([A-Za-z0-9_-]+)/.exec(url);
  if (direct) return direct[1];
  const query = /[?&]id=([A-Za-z0-9_-]+)/.exec(url);
  if (query) return query[1];
  return '';
}

function extractFolderUrl(value: unknown, formula: string) {
  if (formula && /^=HYPERLINK\(/i.test(formula)) {
    const m = /=HYPERLINK\(\s*"([^"]+)"/i.exec(formula);
    if (m) return m[1];
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (/^https?:\/\//i.test(trimmed) || trimmed.includes('drive.google.com')) {
      return trimmed;
    }
  }
  return '';
}

function makeFolderHyperlink(url: string, label = 'Open Folder') {
  if (!url) return '';
  const safeUrl = url.replace(/"/g, '');
  const safeLabel = label.replace(/"/g, '');
  return `=HYPERLINK("${safeUrl}","${safeLabel || 'Open Folder'}")`;
}

const EDITABLE_SONG_COLUMNS = ['Song','Leader','Season','Usage','Themes','Keywords','Scriptures','Notes','Lyrics','Link','Sp','Archive'];

type SaveSongInput = {
  originalName?: string;
  data?: Record<string, unknown>;
};

export function saveSongEntry(input: SaveSongInput) {
  const data = (input?.data && typeof input.data === 'object') ? (input.data as Record<string, unknown>) : {};
  const songName = String(data['Song'] ?? data['song'] ?? '').trim();
  if (!songName) throw new Error('Song title is required.');

  const sh = getSheetByName(SONG_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error('Songs sheet has no columns.');
  if (lastRow < 1) throw new Error('Songs sheet has no header row.');

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const colMap: Record<string, number> = {};
  headers.forEach((h, idx) => {
    const lower = h.toLowerCase();
    colMap[h] = idx;
    colMap[lower] = idx;
  });
  const songIdx = colMap[SONG_COL_NAME] ?? colMap[SONG_COL_NAME.toLowerCase()];
  if (songIdx == null || songIdx < 0) throw new Error('Songs sheet is missing Song column.');

  const norm = (value: unknown) => String(value ?? '').trim().toLowerCase();
  const targetOriginal = norm(input?.originalName);
  const rows = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  let rowOffset = -1;
  if (targetOriginal) {
    rowOffset = rows.findIndex(row => norm(row[songIdx]) === targetOriginal);
  }
  const newNameNorm = norm(songName);
  const duplicateIdx = rows.findIndex((row, idx) => {
    if (rowOffset >= 0 && idx === rowOffset) return false;
    return norm(row[songIdx]) === newNameNorm;
  });
  if (duplicateIdx >= 0) {
    throw new Error('A song with that title already exists.');
  }

  const allowed = new Set(EDITABLE_SONG_COLUMNS.map(label => label.toLowerCase()));
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    let rowNumber: number;
    if (rowOffset >= 0) {
      rowNumber = rowOffset + 2;
    } else {
      rowNumber = lastRow + 1;
      const blankRow = new Array(lastCol).fill('');
      sh.getRange(rowNumber, 1, 1, lastCol).setValues([blankRow]);
    }

    for (const [keyRaw, value] of Object.entries(data)) {
      const key = String(keyRaw || '').trim();
      if (!key) continue;
      const lower = key.toLowerCase();
      if (!allowed.has(lower)) continue;
      const colIndex = colMap[key] ?? colMap[lower];
      if (colIndex == null || colIndex < 0) continue;
      let cellValue: unknown = value ?? '';
      if (typeof cellValue === 'boolean') {
        cellValue = cellValue ? 'Y' : '';
      } else if (typeof cellValue === 'string') {
        const trimmed = cellValue.trim();
        if (lower === 'sp' || lower === 'archive') {
          cellValue = /^y(es)?|true|1$/i.test(trimmed) ? 'Y' : '';
        }
      }
      sh.getRange(rowNumber, colIndex + 1).setValue(cellValue);
    }

    const songColIndex = colMap[SONG_COL_NAME] ?? colMap[SONG_COL_NAME.toLowerCase()];
    if (songColIndex != null && songColIndex >= 0) {
      sh.getRange(rowNumber, songColIndex + 1).setValue(songName);
    }
  } finally {
    lock.releaseLock();
  }

  return { items: getSongsWithLinksForView() };
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

    const songColIdx = colMap[SONG_COL_NAME] ?? -1;
    const spColIdx = colMap[SP_COL_NAME] ?? -1;
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
                if (rt && typeof rt.getLinkUrl === 'function') {
                    const directUrl = rt.getLinkUrl();
                    if (directUrl) folderUrl = directUrl;
                }
                if (!folderUrl && rt && typeof rt.getRuns === 'function') {
                    const runs = rt.getRuns();
                    if (Array.isArray(runs)) {
                        for (const run of runs) {
                            try {
                                const runUrl = typeof run.getLinkUrl === 'function' ? run.getLinkUrl() : null;
                                if (runUrl) { folderUrl = runUrl; break; }
                            } catch (_) { /* ignore */ }
                        }
                    }
                }
            } catch (_) { /* ignore */ }
            if (!folderUrl) {
                const f = formulas[r][fCol];
                const m = /^=HYPERLINK\("([^"]+)"/i.exec(String(f || ''));
                if (m) folderUrl = m[1];
            }
            if (!folderUrl) {
                const raw = String(values[r][fCol] ?? '');
                const match = raw.match(/https:\/\/drive\.google\.com\/[^\s"]+/);
                if (match) folderUrl = match[0];
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
