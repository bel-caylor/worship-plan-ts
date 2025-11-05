import { ROLES_COL, ROLES_SHEET } from '../constants';
import { getSheetByName } from '../util/sheets';

type RolesListItem = {
  email: string;
  permissions: string;
  first: string;
  last: string;
  teams: string[];
  teamRaw: string;
  role: string;
  spanish: string;
};

type UpdateRoleInput = {
  email: string;
  role?: string;
  team?: string;
  permissions?: string;
  spanish?: string | boolean;
};

type AddRoleInput = {
  email: string;
  first: string;
  last: string;
  permissions?: string;
  team?: string;
  role?: string;
  spanish?: string | boolean;
};

const TEAM_SPLIT = /[,;|]/;

export function listRoles() {
  const sh = getSheetByName(ROLES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { items: [], teams: [] as string[] };

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const idxEmail = col(ROLES_COL.email);
  const idxPerms = col(ROLES_COL.permissions);
  const idxFirst = col(ROLES_COL.first);
  const idxLast = col(ROLES_COL.last);
  const idxTeam = col(ROLES_COL.team);
  const idxRole = col(ROLES_COL.role);
  const idxSpanish = col(ROLES_COL.spanish);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const items: RolesListItem[] = [];
  const teamSet = new Set<string>();

  for (const row of body) {
    const email = idxEmail >= 0 ? String(row[idxEmail] ?? '').trim() : '';
    if (!email) continue;
    const teamRaw = idxTeam >= 0 ? String(row[idxTeam] ?? '').trim() : '';
    const teams = teamRaw
      ? teamRaw.split(TEAM_SPLIT).map(s => s.trim()).filter(Boolean)
      : [];
    teams.forEach(t => { if (t) teamSet.add(t); });
    items.push({
      email,
      permissions: idxPerms >= 0 ? String(row[idxPerms] ?? '').trim() : '',
      first: idxFirst >= 0 ? String(row[idxFirst] ?? '').trim() : '',
      last: idxLast >= 0 ? String(row[idxLast] ?? '').trim() : '',
      teams,
      teamRaw,
      role: idxRole >= 0 ? String(row[idxRole] ?? '').trim() : '',
      spanish: idxSpanish >= 0 ? String(row[idxSpanish] ?? '').trim() : ''
    });
  }

  items.sort((a, b) => a.last.localeCompare(b.last) || a.first.localeCompare(b.first));

  return {
    items,
    teams: Array.from(teamSet).sort((a, b) => a.localeCompare(b))
  };
}

export function updateRoleEntry(input: UpdateRoleInput) {
  const email = String(input.email || '').trim().toLowerCase();
  if (!email) throw new Error('Email is required to update roles.');

  const sh = getSheetByName(ROLES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) throw new Error('Roles sheet is empty.');

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => {
    const idx = headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
    if (idx < 0) throw new Error(`Column "${name}" not found on Roles sheet.`);
    return idx;
  };
  const idxEmail = col(ROLES_COL.email);
  const idxTeam = col(ROLES_COL.team);
  const idxRole = col(ROLES_COL.role);
  const idxPerms = col(ROLES_COL.permissions);
  const idxSpanish = col(ROLES_COL.spanish);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let targetIndex = -1;
  for (let i = 0; i < body.length; i++) {
    const rowEmail = String(body[i][idxEmail] ?? '').trim().toLowerCase();
    if (rowEmail === email) {
      targetIndex = i;
      break;
    }
  }
  if (targetIndex < 0) throw new Error(`No role entry found for ${email}`);

  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    const rowNumber = 2 + targetIndex;
    const updates: Array<{ column: number; value: string }> = [];
    if (typeof input.team === 'string') {
      updates.push({ column: idxTeam + 1, value: input.team.trim() });
    }
    if (typeof input.role === 'string') {
      updates.push({ column: idxRole + 1, value: input.role.trim() });
    }
    if (typeof input.permissions === 'string') {
      updates.push({ column: idxPerms + 1, value: input.permissions.trim() });
    }
    if (typeof input.spanish !== 'undefined') {
      let spanishVal = '';
      if (typeof input.spanish === 'boolean') {
        spanishVal = input.spanish ? 'Y' : '';
      } else if (typeof input.spanish === 'string') {
        spanishVal = input.spanish.trim().toUpperCase() === 'Y' ? 'Y' : '';
      }
      updates.push({ column: idxSpanish + 1, value: spanishVal });
    }
    if (!updates.length) return listRoles();
    updates.forEach(u => sh.getRange(rowNumber, u.column).setValue(u.value));
  } finally {
    lock.releaseLock();
  }

  return listRoles();
}

export function addRoleEntry(input: AddRoleInput) {
  const email = String(input.email || '').trim().toLowerCase();
  if (!email) throw new Error('Email is required.');

  const sh = getSheetByName(ROLES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error('Roles sheet has no columns.');

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => {
    const idx = headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
    if (idx < 0) throw new Error(`Column "${name}" not found on Roles sheet.`);
    return idx;
  };
  const idxEmail = col(ROLES_COL.email);
  const idxPerms = col(ROLES_COL.permissions);
  const idxFirst = col(ROLES_COL.first);
  const idxLast = col(ROLES_COL.last);
  const idxTeam = col(ROLES_COL.team);
  const idxRole = col(ROLES_COL.role);
  const idxSpanish = col(ROLES_COL.spanish);

  if (lastRow >= 2) {
    const emails = sh.getRange(2, idxEmail + 1, lastRow - 1, 1).getValues();
    const exists = emails.some(v => String(v[0] ?? '').trim().toLowerCase() === email);
    if (exists) throw new Error(`A member with email ${email} already exists.`);
  }

  const nextRow = lastRow + 1;
  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);
  try {
    if (nextRow === 1) sh.insertRowAfter(1);
    sh.getRange(nextRow, idxEmail + 1).setValue(email);
    sh.getRange(nextRow, idxFirst + 1).setValue(String(input.first || '').trim());
    sh.getRange(nextRow, idxLast + 1).setValue(String(input.last || '').trim());
    if (typeof input.permissions === 'string') {
      sh.getRange(nextRow, idxPerms + 1).setValue(String(input.permissions || '').trim());
    }
    if (typeof input.team === 'string') {
      sh.getRange(nextRow, idxTeam + 1).setValue(String(input.team || '').trim());
    }
    if (typeof input.role === 'string') {
      sh.getRange(nextRow, idxRole + 1).setValue(String(input.role || '').trim());
    }
    if (typeof input.spanish !== 'undefined') {
      let spanishVal = '';
      if (typeof input.spanish === 'boolean') spanishVal = input.spanish ? 'Y' : '';
      else if (typeof input.spanish === 'string') spanishVal = input.spanish.trim().toUpperCase() === 'Y' ? 'Y' : '';
      sh.getRange(nextRow, idxSpanish + 1).setValue(spanishVal);
    }
  } finally {
    lock.releaseLock();
  }

  return listRoles();
}




