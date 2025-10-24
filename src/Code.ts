/** CONFIG **/
const SONG_SHEET = 'Songs';
const SONG_COL_NAME = 'Song';
const FOLDER_LINK_COL = 'Folder URL';
const AUDIO_LINKS_COL = 'Audio Files';
const MAX_AUDIO_LINKS = 5; // limit per row
const AUDIO_MIME_PREFIX = 'audio/'; // fallback check
const AUDIO_EXT = new Set(['mp3', 'm4a', 'aac', 'wav', 'aiff', 'aif', 'flac', 'ogg', 'oga', 'opus', 'wma']);
// Only search inside this shared folder
const ROOT_FOLDER_ID = '19buHshZq5phnvP8xvFHnwkTnwdvg0FV_';
const SPANISH_ROOT_ID = '1bYk1utXCF0D1r5a_GHlAjqmO5gLBg8b5'; // “+Spanish” subfolder
const SP_COL_NAME = 'Sp'; // column header with Y/N

//Cache folders separately per root
let _folderCacheByRoot: Record<string, { id: string; name: string; url: string }[] | null> = {};
function ensureFolderCache(rootId: string) {
    if (!_folderCacheByRoot[rootId]) _folderCacheByRoot[rootId] = listAllSubfolders(rootId);
    return _folderCacheByRoot[rootId]!;
}

// --- Config for Leader indexing ---
const TARGET_LEADER_COL = 'Leader';
const PLANNER_SHEET = 'Weekly Planner';
const PLANNER_LEADER_CANDIDATES = ['Leader'];
const PLANNER_SONG_COLS = [
    'Opening Song',
    'Song2',
    'Song3',
    'Song4/Communion',
    'Offering/Communion Song',
    'Closing Song'
];

// Build unique leaders per song and write to Songs.TARGET_LEADER_COL
function buildLeadersFromPlanner() {
    const songsSh = getSheetByName(SONG_SHEET);
    const plannerSh = getSheetByName(PLANNER_SHEET); // will throw if the tab doesn’t exist

    // --- Read planner header/body
    const pVals = plannerSh.getDataRange().getValues();
    if (pVals.length < 2) {
        SpreadsheetApp.getActive().toast('Planner has no data rows.', 'Weekly Planner', 4);
        return;
    }
    const pHeaders = pVals.shift()!.map(v => String(v ?? '').trim());
    const pLeaderIdx = findHeaderIndex(pHeaders, PLANNER_LEADER_CANDIDATES);
    const pSongIdxs = findManyHeaderIndices(pHeaders, PLANNER_SONG_COLS);

    if (pLeaderIdx < 0) {
        throw new Error(`Could not find a Leader column on "${PLANNER_SHEET}".
            Saw headers: ${pHeaders.join(' | ')}
            Looking for any of: ${PLANNER_LEADER_CANDIDATES.join(', ')}`);
    }
    if (!pSongIdxs.length) {
        throw new Error(`None of the song columns were found on "${PLANNER_SHEET}".
            Looking for: ${PLANNER_SONG_COLS.join(' | ')}`);
    }

    // --- Aggregate leaders per song (case-insensitive key)
    const bySong = new Map<string, Set<string>>();
    for (const row of pVals) {
        const leader = String(row[pLeaderIdx] ?? '').trim();
        if (!leader) continue;

        for (const c of pSongIdxs) {
            const raw = String(row[c] ?? '').trim();
            if (!raw) continue;

            // Split in case the cell contains multiple titles separated by / or ,
            const titles = raw.split(/[\/,;|]/).map(s => s.trim()).filter(Boolean);
            for (const title of titles) {
                const key = title.toLowerCase();
                if (!bySong.has(key)) bySong.set(key, new Set());
                bySong.get(key)!.add(leader);
            }
        }
    }

    // --- Ensure target column exists on Songs
    const { headers: sHeaders, colMap: sColMap } = getHeaders(songsSh);
    if (!(TARGET_LEADER_COL in sColMap)) {
        ensureColumn(songsSh, sHeaders, sColMap, TARGET_LEADER_COL);
    }
    const sSongIdx = sColMap[SONG_COL_NAME];                // index in row arrays
    const outCol = sColMap[TARGET_LEADER_COL] + 1;        // 1-based column number

    const lastRow = songsSh.getLastRow();
    if (lastRow < 2) {
        SpreadsheetApp.getActive().toast('Songs sheet has no data rows.', 'Worship Planner', 4);
        return;
    }

    const body = songsSh.getRange(2, 1, lastRow - 1, songsSh.getLastColumn()).getValues();

    // --- Write the leader list for each song row
    const lock = LockService.getDocumentLock();
    lock.waitLock(10000);
    try {
        for (let i = 0; i < body.length; i++) {
            const songTitle = String(body[i][sSongIdx] ?? '').trim();
            const leaders = songTitle ? bySong.get(songTitle.toLowerCase()) : undefined;
            const text = leaders ? Array.from(leaders).sort((a, b) => a.localeCompare(b)).join(', ') : '';
            songsSh.getRange(2 + i, outCol).setValue(text);
        }
    } finally {
        lock.releaseLock();
    }

    SpreadsheetApp.getActive().toast('Leader list updated from Worship Planner.', 'Worship Planner', 4);
}

type Row = Record<string, unknown>;

/** MENU **/
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Worship')
        .addItem('Link song media', 'linkSongMedia')
        .addItem('Build leader list (from Usage Log)', 'buildLeadersFromPlanner')
        .addToUi();
}

/** MAIN: Scan rows, find folders, add links */
function linkSongMedia() {
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
            const match = findBestFolderForSong(song, rootIdForRow); // <— updated signature
            let folderUrl = '';
            let audioRich: GoogleAppsScript.Spreadsheet.RichTextValue | null = null;

            if (match) {
                folderUrl = match.url;

                // List audio files in that folder
                const audioFiles = listAudioInFolder(match.id, MAX_AUDIO_LINKS);
                if (audioFiles.length) {
                    const builder = SpreadsheetApp.newRichTextValue();
                    let text = '';
                    const runs: { start: number, end: number, url: string }[] = [];
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

            // Folder URL: use a HYPERLINK formula so it’s clickable with a friendly label
            if (folderUrl) {
                const label = 'Open Folder';
                sh.getRange(rowIndex, folderCol).setFormula(`=HYPERLINK("${folderUrl}","${label}")`);
            } else {
                sh.getRange(rowIndex, folderCol).clearContent();
            }

            // Audio links: rich text with one link per line (filenames)
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

/** Helpers **/
function getHeaders(sh: GoogleAppsScript.Spreadsheet.Sheet) {
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const colMap: Record<string, number> = {};
    headers.forEach((h, i) => (colMap[h] = i));
    if (!(SONG_COL_NAME in colMap)) throw new Error(`Header "${SONG_COL_NAME}" not found`);
    return { headers, colMap };
}

function ensureColumn(
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

/** Find the best folder using DriveApp search + fuzzy match */
// Cache folder list in memory while script runs
let _folderCache: { id: string; name: string; url: string }[] | null = null;

/** Find best folder by fuzzy match within ROOT_FOLDER_ID */
function findBestFolderForSong(songName: string, rootId: string) {
    const folders = ensureFolderCache(rootId);

    const normSong = normalize(songName);
    const exact = folders.find(f => normalize(f.name) === normSong);
    if (exact) return { ...exact, score: 1 };

    const contains = folders.find(f => normalize(f.name).includes(normSong));
    if (contains) return { ...contains, score: 0.9 };

    let best: { id: string; name: string; url: string; score: number } | null = null;
    for (const f of folders) {
        const s = similarity(normalize(f.name), normSong);
        if (!best || s > best.score) best = { ...f, score: s };
    }
    return best && best.score >= 0.35 ? best : null;
}

/** List all files in a folder URL for preview (name, url, mimeType). */
function getFilesForFolderUrl(folderUrl: string, limit: number = 200) {
    if (!folderUrl) return [];
    const m = /\/folders\/([A-Za-z0-9_-]+)/.exec(folderUrl);
    if (!m) return [];
    const folder = DriveApp.getFolderById(m[1]);
    const out: Array<{ name: string; url: string; mimeType: string }> = [];
    const it = folder.getFiles();
    while (it.hasNext() && out.length < limit) {
        const f = it.next();
        out.push({ name: f.getName(), url: f.getUrl(), mimeType: f.getMimeType() });
    }
    // Optionally sort by name
    out.sort((a, b) => a.name.localeCompare(b.name));
    return out;
}

/** Recursively list all subfolders under a root folder */
function listAllSubfolders(rootId: string) {
    const root = DriveApp.getFolderById(rootId);
    const stack = [root];
    const all: { id: string; name: string; url: string }[] = [];

    while (stack.length) {
        const folder = stack.pop()!;
        const subfolders = folder.getFolders();
        while (subfolders.hasNext()) {
            const f = subfolders.next();
            all.push({ id: f.getId(), name: f.getName(), url: f.getUrl() });
            stack.push(f);
        }
    }
    return all;
}


/** List audio files from a folder (by id) */
function listAudioInFolder(folderId: string, limit: number) {
    const folder = DriveApp.getFolderById(folderId);
    const out: { name: string; url: string }[] = [];
    const files = folder.getFiles();
    while (files.hasNext() && out.length < limit) {
        const file = files.next();
        const mime = file.getMimeType() || '';
        const name = file.getName();
        const ext = (name.split('.').pop() || '').toLowerCase();
        if (mime.startsWith(AUDIO_MIME_PREFIX) || AUDIO_EXT.has(ext)) {
            out.push({ name, url: file.getUrl() });
        }
    }
    return out;
}

/** Build rows for the UI and include parsed link info (robust). */
function getSongsWithLinksForView(): Row[] {
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
            : (colMap['Leader'] ?? -1);                // <-- NEW (supports 'Leader' or your TARGET_LEADER_COL)

    const bodyRange = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn());
    const values = bodyRange.getValues();
    const formulas = bodyRange.getFormulas();
    const rich = bodyRange.getRichTextValues();

    const out: Row[] = [];

    for (let r = 0; r < values.length; r++) {
        const rowObj: Row = {};
        headers.forEach((h, c) => (rowObj[h] = values[r][c]));

        // ---- Folder URL (rich text or HYPERLINK fallback) ----
        let folderUrl: string | null = null;
        if (fCol >= 0) {
            const rt = rich[r][fCol];
            try {
                if (rt && typeof (rt as GoogleAppsScript.Spreadsheet.RichTextValue).getLinkUrl === 'function') {
                    const url = (rt as GoogleAppsScript.Spreadsheet.RichTextValue).getLinkUrl();
                    if (url) folderUrl = url;
                }
            } catch { }
            if (!folderUrl) {
                const f = formulas[r][fCol];
                const m = /^=HYPERLINK\("([^"]+)"/i.exec(String(f || ''));
                if (m) folderUrl = m[1];
            }
        }
        (rowObj as any)._folderUrl = folderUrl || null;

        // ---- Leaders array for filtering (split on commas/slashes/etc.) ----
        const leadersRaw =
            leaderColIdx != null && leaderColIdx >= 0 ? String(values[r][leaderColIdx] ?? '') : '';  // <-- NEW
        (rowObj as any)._leaders = splitTokens(leadersRaw);                                          // <-- NEW

        // optional placeholder, since we list media live
        (rowObj as any)._audio = [];

        out.push(rowObj);
    }

    return out;
}


/** Utils: similarity & normalization */

// Case/space-insensitive normalize
const norm = (s: string) => String(s).toLowerCase().replace(/\s+/g, '');

// Find one column by trying multiple names
function findHeaderIndex(headers: string[], candidates: string[]): number {
    const H = headers.map(norm);
    for (const cand of candidates) {
        const i = H.indexOf(norm(cand));
        if (i >= 0) return i;
    }
    return -1;
}

// Find several columns by their exact labels (tolerant to spaces/case)
function findManyHeaderIndices(headers: string[], labels: string[]): number[] {
    const H = headers.map(norm);
    const idxs: number[] = [];
    for (const label of labels) {
        const i = H.indexOf(norm(label));
        if (i >= 0) idxs.push(i);
    }
    return idxs;
}


function splitTokens(s: string) {
    return String(s || '')
        .split(/[\/,;|&]|,\s*/g)   // split on common separators
        .map(t => t.trim())
        .filter(Boolean);
}


function normalize(s: string) {
    return s
        .toLowerCase()
        .replace(/[\p{P}\p{S}]+/gu, ' ') // remove punctuation/symbols
        .replace(/\s+/g, ' ')
        .trim();
}

function escapeQuery(s: string) {
    return s.replace(/"/g, '\\"');
}

// Normalized Levenshtein similarity (0..1)
function similarity(a: string, b: string) {
    const d = levenshtein(a, b);
    const maxLen = Math.max(a.length, b.length) || 1;
    return 1 - d / maxLen;
}

function levenshtein(a: string, b: string) {
    const m = a.length, n = b.length;
    if (m === 0) return n;
    if (n === 0) return m;

    const dp = new Array(n + 1);
    for (let j = 0; j <= n; j++) dp[j] = j;

    for (let i = 1; i <= m; i++) {
        let prev = dp[0];
        dp[0] = i;
        for (let j = 1; j <= n; j++) {
            const tmp = dp[j];
            dp[j] = Math.min(
                dp[j] + 1,              // deletion
                dp[j - 1] + 1,          // insertion
                prev + (a[i - 1] === b[j - 1] ? 0 : 1) // substitution
            );
            prev = tmp;
        }
    }
    return dp[n];
}


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
    if (e?.parameter?.json === '1') {
        return ContentService.createTextOutput(JSON.stringify({ rows: getSongsWithLinksForView() }))
            .setMimeType(ContentService.MimeType.JSON);
    }
    const tpl = HtmlService.createTemplateFromFile('html/index');
    tpl.rowsJson = JSON.stringify(getSongsWithLinksForView());   // inject with links
    return tpl.evaluate().setTitle('Worship Planner');
}


