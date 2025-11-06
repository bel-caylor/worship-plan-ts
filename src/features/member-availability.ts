import { MEMBER_AVAILABILITY_COL, MEMBER_AVAILABILITY_SHEET } from '../constants';
import { getSheetByName } from '../util/sheets';

type AvailabilityPayload = {
  email: string;
  unavailableServiceIds?: string[];
  serviceIds?: string[]; // legacy alias
  statusLabel?: string;
};

const normalizeEmail = (value: any) => String(value ?? '').trim().toLowerCase();

function getHeaderIndexes(sh: GoogleAppsScript.Spreadsheet.Sheet) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const findIdx = (name: string) => {
    const idx = headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
    if (idx === -1) throw new Error(`Column "${name}" not found in ${MEMBER_AVAILABILITY_SHEET}`);
    return idx;
  };
  const emailIdx = findIdx(MEMBER_AVAILABILITY_COL.email);
  const serviceIdIdx = findIdx(MEMBER_AVAILABILITY_COL.serviceId);
  const availabilityIdx = findIdx(MEMBER_AVAILABILITY_COL.availability);
  return { emailIdx, serviceIdIdx, availabilityIdx, lastCol };
}

export function getMemberAvailability(payload: { email: string }) {
  const email = normalizeEmail(payload?.email);
  if (!email) throw new Error('Email is required');
  const sh = getSheetByName(MEMBER_AVAILABILITY_SHEET);
  const { emailIdx, serviceIdIdx, availabilityIdx, lastCol } = getHeaderIndexes(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2 || lastCol < 1) return { unavailable: [] };
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const unavailable: string[] = [];
  for (const row of values) {
    const rowEmail = normalizeEmail(row[emailIdx]);
    if (rowEmail !== email) continue;
    const serviceId = String(row[serviceIdIdx] ?? '').trim();
    if (!serviceId) continue;
    const availability = String(row[availabilityIdx] ?? '').trim().toLowerCase();
    if (!availability || availability.startsWith('u')) {
      unavailable.push(serviceId);
    }
  }
  return { unavailable };
}

export function saveMemberAvailability(payload: AvailabilityPayload) {
  const email = normalizeEmail(payload?.email);
  if (!email) throw new Error('Email is required');
  const idsRaw = Array.isArray(payload?.unavailableServiceIds)
    ? payload?.unavailableServiceIds
    : payload?.serviceIds;
  const unavailableIds = Array.isArray(idsRaw)
    ? idsRaw.map(id => String(id ?? '').trim()).filter(Boolean)
    : [];
  const sh = getSheetByName(MEMBER_AVAILABILITY_SHEET);
  const { emailIdx, serviceIdIdx, availabilityIdx, lastCol } = getHeaderIndexes(sh);
  const label = String(payload?.statusLabel || 'Unavailable');

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
      const matches: number[] = [];
      values.forEach((row, idx) => {
        if (normalizeEmail(row[emailIdx]) === email) {
          matches.push(idx + 2); // account for header row
        }
      });
      matches.sort((a, b) => b - a).forEach(rowNum => {
        sh.deleteRow(rowNum);
      });
    }

    if (unavailableIds.length) {
      const rows = unavailableIds.map(serviceId => {
        const cols = Math.max(lastCol, sh.getLastColumn());
        const vals = Array.from({ length: cols }, () => '');
        vals[emailIdx] = email;
        vals[serviceIdIdx] = serviceId;
        vals[availabilityIdx] = label;
        return vals;
      });
      const startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    }
  } finally {
    lock.releaseLock();
  }

  return { count: unavailableIds.length };
}
