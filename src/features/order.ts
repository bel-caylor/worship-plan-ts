// src/features/order.ts
import { ORDER_SHEET, ORDER_COL } from '../constants';
import { getSheetByName } from '../util/sheets';
import { songUsageForItemType, updateSongRecency } from './songs';
import { getLatestSongPerformancePlaybacks } from './services';

const ORDER_CACHE_PREFIX = 'wp.order.v1:';
const ORDER_CACHE_TTL_SECONDS = 300;

export type OrderItem = {
  order: number;
  itemType: string;
  detail?: string;
  scriptureText?: string;
  leader?: string;
  notes?: string;
  recordingUrl?: string;
  recordingLabel?: string;
};

export type OrderRecordingLink = {
  songName: string;
  recordingUrl: string;
  recordingLabel: string;
};

const orderCacheKey = (serviceId: string) => `${ORDER_CACHE_PREFIX}${encodeURIComponent(serviceId)}`;

function readOrderCache(serviceId: string): { items: OrderItem[] } | null {
  try {
    const raw = CacheService.getDocumentCache().get(orderCacheKey(serviceId));
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed?.items) ? { items: parsed.items as OrderItem[] } : null;
  } catch (_) {
    return null;
  }
}

function writeOrderCache(serviceId: string, items: OrderItem[]) {
  try {
    CacheService.getDocumentCache().put(
      orderCacheKey(serviceId),
      JSON.stringify({ items }),
      ORDER_CACHE_TTL_SECONDS
    );
  } catch (_) {
    // Cache capacity is limited; a cache miss must never affect order loading.
  }
}

function orderItemForCache(item: OrderItem, fallbackOrder: number): OrderItem {
  return {
    order: Number(item?.order ?? fallbackOrder),
    itemType: String(item?.itemType ?? ''),
    detail: String(item?.detail ?? ''),
    scriptureText: String(item?.scriptureText ?? ''),
    leader: String(item?.leader ?? ''),
    notes: String(item?.notes ?? '')
  };
}

export function getOrder(serviceId: string) {
  const sid = String(serviceId || '').trim();
  if (!sid) return { items: [] };
  const cached = readOrderCache(sid);
  if (cached) return cached;
  const sh = getSheetByName(ORDER_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    const empty = { items: [] as OrderItem[] };
    writeOrderCache(sid, empty.items);
    return empty;
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const serviceIdx = col(ORDER_COL.serviceId);
  const orderIdx = col(ORDER_COL.order);
  const typeIdx = col(ORDER_COL.itemType);
  const detailIdx = col(ORDER_COL.detail);
  const scriptureTextIdx = col(ORDER_COL.scriptureText);
  const leaderIdx = col(ORDER_COL.leader);
  const notesIdx = col(ORDER_COL.notes);
  if (serviceIdx < 0) {
    const empty = { items: [] as OrderItem[] };
    writeOrderCache(sid, empty.items);
    return empty;
  }

  // The historical implementation read every cell in the Order sheet before
  // filtering. Scripture text can make that payload quite large. Locate the
  // service using one narrow column, then read only its contiguous row block(s).
  const serviceIds = sh.getRange(2, serviceIdx + 1, lastRow - 1, 1).getDisplayValues();
  const matchingRows = serviceIds
    .map((row, index) => String(row[0] ?? '').trim() === sid ? index + 2 : 0)
    .filter(Boolean);
  if (!matchingRows.length) {
    const empty = { items: [] as OrderItem[] };
    writeOrderCache(sid, empty.items);
    return empty;
  }

  const rowBlocks: Array<{ start: number; count: number }> = [];
  matchingRows.forEach((rowNumber) => {
    const previous = rowBlocks[rowBlocks.length - 1];
    if (previous && previous.start + previous.count === rowNumber) {
      previous.count += 1;
    } else {
      rowBlocks.push({ start: rowNumber, count: 1 });
    }
  });
  const items: OrderItem[] = [];
  rowBlocks.forEach(({ start, count }) => {
    sh.getRange(start, 1, count, lastCol).getValues().forEach((row) => {
      const detail = detailIdx >= 0 ? String(row[detailIdx] ?? '') : '';
      items.push({
        order: orderIdx >= 0 ? Number(row[orderIdx] ?? 0) : 0,
        itemType: typeIdx >= 0 ? String(row[typeIdx] ?? '') : '',
        detail,
        scriptureText: scriptureTextIdx >= 0 ? String(row[scriptureTextIdx] ?? '') : '',
        leader: leaderIdx >= 0 ? String(row[leaderIdx] ?? '') : '',
        notes: notesIdx >= 0 ? String(row[notesIdx] ?? '') : ''
      });
    });
  });
  items.sort((a, b) => a.order - b.order);
  writeOrderCache(sid, items);
  return { items };
}

/**
 * Recording history is helpful context, but it must not delay rendering the
 * saved order. The client requests these links after it has painted the order.
 */
export function getOrderRecordingLinks(input: { songNames?: string[] }) {
  const songNames = Array.from(new Set(
    (Array.isArray(input?.songNames) ? input.songNames : [])
      .map(name => String(name || '').trim())
      .filter(Boolean)
  ));
  const playbackBySong = getLatestSongPerformancePlaybacks(songNames);
  const items: OrderRecordingLink[] = songNames.map(songName => {
    const playback = playbackBySong.get(songName);
    const hasTimestamp = Number(playback?.startSeconds || 0) > 0;
    return {
      songName,
      recordingUrl: hasTimestamp ? String(playback?.youtubeUrl || '') : '',
      recordingLabel: hasTimestamp
        ? (playback?.startLabel ? `Open at ${playback.startLabel}` : 'Open recording')
        : ''
    };
  });
  return { items };
}

export function saveOrder(input: { serviceId: string; items: OrderItem[]; serviceDate?: string }) {
  const serviceId = String(input?.serviceId || '').trim();
  const items = Array.isArray(input?.items) ? input.items : [];
  // The service ID is immutable once a service exists, while the client form can
  // briefly contain a stale date during a service switch.  Use the ID's date so
  // an autosave cannot stamp songs with another service's date.
  const serviceDate = dateFromServiceId(serviceId) || normalizeServiceDate(input?.serviceDate);
  if (!serviceId) throw new Error('serviceId required');
  const sh = getSheetByName(ORDER_SHEET);

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const serviceIdx = col(ORDER_COL.serviceId);
  const orderIdx = col(ORDER_COL.order);
  const typeIdx = col(ORDER_COL.itemType);
  const detailIdx = col(ORDER_COL.detail);
  const scriptureTextIdx = col(ORDER_COL.scriptureText);
  const leaderIdx = col(ORDER_COL.leader);
  const notesIdx = col(ORDER_COL.notes);

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    // Efficient in-place update: reuse existing rows for this service where possible
    const lastRow = sh.getLastRow();
    const existing: { sheetRow: number; order: number }[] = [];
    const byOrder = new Map<number, number>(); // order -> sheetRow
    if (lastRow >= 2 && serviceIdx >= 0) {
      const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      for (let i = 0; i < body.length; i++) {
        const row = body[i];
        const sid = String(row[serviceIdx] ?? '').trim();
        if (sid !== serviceId) continue;
        const ord = orderIdx >= 0 ? Number(row[orderIdx] ?? 0) : 0;
        const sheetRow = 2 + i;
        existing.push({ sheetRow, order: ord });
        if (!isNaN(ord) && ord > 0 && !byOrder.has(ord)) byOrder.set(ord, sheetRow);
      }
    }

    const unused = new Set(existing.map(e => e.sheetRow));
    const pickVals = (it: OrderItem, idx: number) => {
      const vals: any[] = Array.from({ length: lastCol }, () => '');
      if (serviceIdx >= 0) vals[serviceIdx] = serviceId;
      if (orderIdx >= 0) vals[orderIdx] = Number(it.order ?? idx + 1);
      if (typeIdx >= 0) vals[typeIdx] = it.itemType ?? '';
      if (detailIdx >= 0) vals[detailIdx] = it.detail ?? '';
      if (scriptureTextIdx >= 0) vals[scriptureTextIdx] = it.scriptureText ?? '';
      if (leaderIdx >= 0) vals[leaderIdx] = it.leader ?? '';
      if (notesIdx >= 0) vals[notesIdx] = it.notes ?? '';
      return vals;
    };

    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const desiredOrder = Number(it.order ?? i + 1);
      let targetRow = byOrder.get(desiredOrder) || null;
      if (!targetRow) {
        // reuse any unused existing row for this service
        const firstUnused = Array.from(unused.values())[0];
        if (firstUnused) targetRow = firstUnused;
      }
      const vals = pickVals(it, i);
      if (targetRow) {
        sh.getRange(targetRow, 1, 1, lastCol).setValues([vals]);
        unused.delete(targetRow);
      } else {
        // append if none to reuse
        sh.appendRow(vals);
      }
    }

    // Remove any leftover rows for this service (extras)
    const toDelete = Array.from(unused.values()).sort((a, b) => b - a);
    for (const r of toDelete) sh.deleteRow(r);
  } finally {
    lock.releaseLock();
  }

  // Keep the next read for this service off the full Order-sheet scan. This
  // cache is also the invalidation point for every planner autosave.
  writeOrderCache(serviceId, items.map((item, index) => orderItemForCache(item, index + 1)));

  try {
    updateSongsFromOrder(items, serviceDate);
  } catch (err) {
    try { Logger.log(`updateSongsFromOrder failed: ${err}`); } catch (_) { }
  }

  return { ok: true };
}

function normalizeServiceDate(input?: string) {
  const raw = String(input || '').trim();
  if (!raw) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  try {
    const d = new Date(raw);
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    }
  } catch (_) { /* ignore */ }
  return raw;
}

function dateFromServiceId(serviceId: string) {
  const match = String(serviceId || '').match(/^(\d{4}-\d{2}-\d{2})(?:_|\b)/);
  return match ? match[1] : '';
}

function looksLikeSongSlot(label: string) {
  return Boolean(songUsageForItemType(label));
}

function updateSongsFromOrder(items: OrderItem[], serviceDate?: string) {
  if (!Array.isArray(items) || !items.length) return;
  const seen = new Set<string>();
  const date = String(serviceDate || '').trim();
  for (const it of items) {
    const usageLabel = songUsageForItemType(String(it?.itemType || ''));
    const detail = String(it?.detail || '').trim();
    if (!detail || !usageLabel) continue;
    const leader = String(it?.leader || '').trim();
    const leaderKey = leader.toLowerCase();
    const key = `${detail.toLowerCase()}|${usageLabel.toLowerCase()}|${leaderKey}`;
    if (seen.has(key)) continue;
    seen.add(key);
    const updateInput: any = {
      name: detail,
      usage: usageLabel,
      incrementUses: false
    };
    if (leader) updateInput.leader = leader;
    if (date) updateInput.date = date;
    try {
      updateSongRecency(updateInput);
    } catch (err) {
      try { Logger.log(`updateSongRecency failed for ${detail}: ${err}`); } catch (_) { }
    }
  }
}
