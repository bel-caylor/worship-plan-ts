// src/features/songs.ts
import {
  SONG_SHEET, SONG_COL_NAME, FOLDER_LINK_COL, AUDIO_LINKS_COL, MAX_AUDIO_LINKS,
  ROOT_FOLDER_ID, SPANISH_ROOT_ID, SP_COL_NAME, TARGET_LEADER_COL, Row
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
