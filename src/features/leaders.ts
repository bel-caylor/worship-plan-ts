// src/features/leaders.ts
import {
  ORDER_COL, ORDER_SHEET, SERVICES_COL, SERVICES_SHEET,
  SONG_COL_NAME, SONG_SHEET, TARGET_LEADER_COL
} from '../constants';
import { songUsageForItemType } from './songs';
import { ensureColumn, findHeaderIndex, getHeaders, getSheetByName } from '../util/sheets';

/** Rebuild the Songs Leader column from the current Services and ServiceItems sheets. */
export function buildLeadersFromPlanner() {
  const servicesSh = getSheetByName(SERVICES_SHEET);
  const itemsSh = getSheetByName(ORDER_SHEET);
  const songsSh = getSheetByName(SONG_SHEET);

  const serviceValues = servicesSh.getDataRange().getValues();
  const serviceHeaders = (serviceValues.shift() || []).map(v => String(v ?? '').trim());
  const serviceIdIdx = findHeaderIndex(serviceHeaders, [SERVICES_COL.id]);
  const serviceLeaderIdx = findHeaderIndex(serviceHeaders, [SERVICES_COL.leader]);
  if (serviceIdIdx < 0 || serviceLeaderIdx < 0) {
    throw new Error(`Services must include "${SERVICES_COL.id}" and "${SERVICES_COL.leader}" columns.`);
  }

  const serviceLeaders = new Map<string, string>();
  for (const row of serviceValues) {
    const id = String(row[serviceIdIdx] ?? '').trim();
    const leader = String(row[serviceLeaderIdx] ?? '').trim();
    if (id && leader) serviceLeaders.set(id, leader);
  }

  const itemValues = itemsSh.getDataRange().getValues();
  const itemHeaders = (itemValues.shift() || []).map(v => String(v ?? '').trim());
  const itemServiceIdIdx = findHeaderIndex(itemHeaders, [ORDER_COL.serviceId]);
  const itemTypeIdx = findHeaderIndex(itemHeaders, [ORDER_COL.itemType]);
  const itemDetailIdx = findHeaderIndex(itemHeaders, [ORDER_COL.detail]);
  const itemLeaderIdx = findHeaderIndex(itemHeaders, [ORDER_COL.leader]);
  if (itemServiceIdIdx < 0 || itemTypeIdx < 0 || itemDetailIdx < 0) {
    throw new Error(`ServiceItems must include "${ORDER_COL.serviceId}", "${ORDER_COL.itemType}", and "${ORDER_COL.detail}" columns.`);
  }

  const bySong = new Map<string, Set<string>>();
  for (const row of itemValues) {
    if (!songUsageForItemType(String(row[itemTypeIdx] ?? ''))) continue;
    const song = String(row[itemDetailIdx] ?? '').trim();
    const serviceId = String(row[itemServiceIdIdx] ?? '').trim();
    const itemLeader = itemLeaderIdx >= 0 ? String(row[itemLeaderIdx] ?? '').trim() : '';
    const leader = itemLeader || serviceLeaders.get(serviceId) || '';
    if (!song || !leader) continue;
    const key = song.toLowerCase();
    if (!bySong.has(key)) bySong.set(key, new Set());
    bySong.get(key)!.add(leader);
  }

  const { headers, colMap } = getHeaders(songsSh);
  ensureColumn(songsSh, headers, colMap, TARGET_LEADER_COL);
  const songIdx = colMap[SONG_COL_NAME];
  const leaderCol = colMap[TARGET_LEADER_COL];
  const lastRow = songsSh.getLastRow();
  if (lastRow < 2) return { updated: 0 };

  const range = songsSh.getRange(2, 1, lastRow - 1, songsSh.getLastColumn());
  const values = range.getValues();
  let updated = 0;
  for (const row of values) {
    const song = String(row[songIdx] ?? '').trim();
    const leaders = bySong.get(song.toLowerCase());
    const value = leaders ? Array.from(leaders).sort((a, b) => a.localeCompare(b)).join(', ') : '';
    if (String(row[leaderCol] ?? '') !== value) {
      row[leaderCol] = value;
      updated++;
    }
  }
  range.setValues(values);
  SpreadsheetApp.getActive().toast(`Leader list rebuilt from Services and ServiceItems (${updated} updated).`, 'Worship Planner', 4);
  return { updated };
}
