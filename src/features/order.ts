// src/features/order.ts
import { ORDER_SHEET, ORDER_COL } from '../constants';
import { getSheetByName } from '../util/sheets';

export type OrderItem = {
  order: number;
  itemType: string;
  detail?: string;
  leader?: string;
  notes?: string;
};

export function getOrder(serviceId: string) {
  const sid = String(serviceId || '').trim();
  if (!sid) return { items: [] };
  const sh = getSheetByName(ORDER_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { items: [] };

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const serviceIdx = col(ORDER_COL.serviceId);
  const orderIdx = col(ORDER_COL.order);
  const typeIdx = col(ORDER_COL.itemType);
  const detailIdx = col(ORDER_COL.detail);
  const scriptureTextIdx = col(ORDER_COL.scriptureText);
  const leaderIdx = col(ORDER_COL.leader);
  const notesIdx = col(ORDER_COL.notes);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const items: OrderItem[] = [];
  for (const r of body) {
    const id = serviceIdx >= 0 ? String(r[serviceIdx] ?? '').trim() : '';
    if (id !== sid) continue;
    items.push({
      order: orderIdx >= 0 ? Number(r[orderIdx] ?? 0) : 0,
      itemType: typeIdx >= 0 ? String(r[typeIdx] ?? '') : '',
      detail: detailIdx >= 0 ? String(r[detailIdx] ?? '') : '',
      // @ts-ignore include optional field
      scriptureText: scriptureTextIdx >= 0 ? String(r[scriptureTextIdx] ?? '') : '',
      leader: leaderIdx >= 0 ? String(r[leaderIdx] ?? '') : '',
      notes: notesIdx >= 0 ? String(r[notesIdx] ?? '') : ''
    });
  }
  items.sort((a, b) => a.order - b.order);
  return { items };
}

export function saveOrder(input: { serviceId: string; items: OrderItem[] }) {
  const serviceId = String(input?.serviceId || '').trim();
  const items = Array.isArray(input?.items) ? input.items : [];
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
    const pickVals = (it: OrderItem & { scriptureText?: string }, idx: number) => {
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
  return { ok: true };
}
