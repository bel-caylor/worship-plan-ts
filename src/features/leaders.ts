// src/features/leaders.ts
import {
  TARGET_LEADER_COL, SONG_SHEET, PLANNER_SHEET, PLANNER_LEADER_CANDIDATES,
  PLANNER_SONG_COLS, SONG_COL_NAME
} from '../constants';
import { getSheetByName, getHeaders, ensureColumn, findHeaderIndex, findManyHeaderIndices } from '../util/sheets';

export function buildLeadersFromPlanner() {
    const songsSh = getSheetByName(SONG_SHEET);
    const plannerSh = getSheetByName(PLANNER_SHEET);

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
        throw new Error(
            `Could not find a Leader column on "${PLANNER_SHEET}".\nSaw: ${pHeaders.join(' | ')}\nLooking for any of: ${PLANNER_LEADER_CANDIDATES.join(', ')}`
        );
    }
    if (!pSongIdxs.length) {
        throw new Error(`None of the song columns were found on "${PLANNER_SHEET}". Looking for: ${PLANNER_SONG_COLS.join(' | ')}`);
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
    const sSongIdx = sColMap[SONG_COL_NAME];
    const outCol = sColMap[TARGET_LEADER_COL] + 1;

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
