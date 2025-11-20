// src/features/song-ai.ts
import { SONG_SHEET, SONG_COL_NAME, TARGET_LEADER_COL } from '../constants';
import { getSheetByName, getHeaders, findHeaderIndex } from '../util/sheets';

export type SongLyricSample = {
  name: string;
  lyrics: string;
};

type SongWithMeta = SongLyricSample & {
  leaderMatch: boolean;
  leaderKey: string;
  lastUsedTs: number;
};

export function loadSongLyricSamples(limit = 24, maxChars = 420, preferredLeader?: string): SongLyricSample[] {
  const sh = getSheetByName(SONG_SHEET);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sh.getLastColumn();
  const { headers, colMap } = getHeaders(sh);
  const nameIdx = colMap[SONG_COL_NAME];
  const lyricsIdx = findHeaderIndex(headers, ['Lyrics', 'Lyric', 'Text']);
  if (nameIdx == null || lyricsIdx < 0) return [];
  const leaderIdx = colMap[TARGET_LEADER_COL] ?? findHeaderIndex(headers, [TARGET_LEADER_COL, 'Target Leader']);
  const lastUsedIdx = findHeaderIndex(headers, ['Last Used', 'Last_Used', 'Last used', 'Last']);

  const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const seen = new Set<string>();
  const normalizedPreferredLeader = normalizeLeader(preferredLeader);
  const samples: SongWithMeta[] = [];

  for (const row of rows) {
    const name = String(row[nameIdx] ?? '').trim();
    const lyricsRaw = String(row[lyricsIdx] ?? '').trim();
    if (!name || !lyricsRaw) continue;
    const dedupeKey = name.toLowerCase();
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);
    const flattened = lyricsRaw.replace(/\s+/g, ' ').trim();
    if (!flattened) continue;
    const snippet = flattened.length > maxChars ? `${flattened.slice(0, maxChars)}...` : flattened;
    const leaderRaw = leaderIdx != null && leaderIdx >= 0 ? row[leaderIdx] : '';
    const leaderKey = normalizeLeader(leaderRaw);
    const leaderMatch = !!normalizedPreferredLeader && !!leaderKey && leaderKey === normalizedPreferredLeader;
    const lastUsedRaw = lastUsedIdx >= 0 ? row[lastUsedIdx] : '';
    const lastUsedTs = parseLastUsed(lastUsedRaw);
    samples.push({ name, lyrics: snippet, leaderMatch, leaderKey, lastUsedTs });
  }

  samples.sort((a, b) => {
    if (a.leaderMatch !== b.leaderMatch) return a.leaderMatch ? -1 : 1;
    if (a.leaderKey !== b.leaderKey) return a.leaderKey.localeCompare(b.leaderKey);
    if (a.lastUsedTs !== b.lastUsedTs) return a.lastUsedTs - b.lastUsedTs;
    return a.name.localeCompare(b.name);
  });

  return samples.slice(0, limit).map(({ name, lyrics }) => ({ name, lyrics }));
}

function normalizeLeader(value: unknown) {
  return String(value || '').trim().toLowerCase();
}

function parseLastUsed(value: unknown): number {
  if (value instanceof Date) {
    const ts = value.getTime();
    return isNaN(ts) ? 0 : ts;
  }
  if (typeof value === 'number' && isFinite(value)) {
    const ts = new Date(value).getTime();
    if (!isNaN(ts)) return ts;
  }
  const str = String(value || '').trim();
  if (!str) return 0;
  const parsed = Date.parse(str);
  return isNaN(parsed) ? 0 : parsed;
}
