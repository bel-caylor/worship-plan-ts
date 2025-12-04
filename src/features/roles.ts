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

type ViewerCapabilities = {
  canViewPlan: boolean;
  canEditPlan: boolean;
  canEditSongs: boolean;
  canViewTeam: boolean;
  canManageTeams: boolean;
  canAdminAvailability: boolean;
};

type ViewerProfile = {
  email: string;
  permissions: string;
  first: string;
  last: string;
  isAdmin: boolean;
  isLoggedIn: boolean;
  capabilities: ViewerCapabilities;
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
const normalizeEmail = (value: unknown) => String(value ?? '').trim().toLowerCase();
const normalizePermission = (value: unknown) => String(value ?? '').trim().toLowerCase();
const canonicalizeRoleLabel = (value: unknown) => {
  const role = String(value ?? '').trim();
  if (!role) return '';
  return /^vocals?$/i.test(role) ? 'Vocal' : role;
};

const emptyCapabilities: ViewerCapabilities = {
  canViewPlan: false,
  canEditPlan: false,
  canEditSongs: false,
  canViewTeam: false,
  canManageTeams: false,
  canAdminAvailability: false
};

function computeCapabilities(permission: string | undefined | null) {
  const normalized = normalizePermission(permission);
  const isAdmin = normalized === 'administrator' || normalized === 'admin';
  const isEditor = normalized === 'editor';
  const isSubscriber = normalized === 'subscriber';
  const capabilities: ViewerCapabilities = {
    canViewPlan: isAdmin || isEditor || isSubscriber,
    canEditPlan: isAdmin || isEditor,
    canEditSongs: isAdmin || isEditor,
    canViewTeam: isAdmin,
    canManageTeams: isAdmin,
    canAdminAvailability: isAdmin || isEditor
  };
  return { capabilities, isAdmin };
}

function acquireRolesLock(maxAttempts = 4, waitMs = 5000) {
  const baseDelay = 200;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const lock = LockService.getDocumentLock();
    if (lock.tryLock(waitMs)) {
      return lock;
    }
    Utilities.sleep(baseDelay * Math.pow(2, attempt));
  }
  const finalLock = LockService.getDocumentLock();
  finalLock.waitLock(waitMs);
  return finalLock;
}

export function memberExistsInRoles(input: { email: string }) {
  const email = normalizeEmail(input?.email);
  if (!email) throw new Error('Email is required.');

  const sh = getSheetByName(ROLES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { exists: false };

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const idxEmail = col(ROLES_COL.email);
  if (idxEmail < 0) throw new Error(`Column "${ROLES_COL.email}" not found on Roles sheet.`);
  const idxFirst = col(ROLES_COL.first);
  const idxLast = col(ROLES_COL.last);
  const idxTeam = col(ROLES_COL.team);
  const idxRole = col(ROLES_COL.role);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (const row of body) {
    const rowEmail = normalizeEmail(row[idxEmail]);
    if (rowEmail !== email) continue;
    const teamRaw = idxTeam >= 0 ? String(row[idxTeam] ?? '').trim() : '';
    const teams = teamRaw ? teamRaw.split(TEAM_SPLIT).map(s => s.trim()).filter(Boolean) : [];
    return {
      exists: true,
      member: {
        email,
        first: idxFirst >= 0 ? String(row[idxFirst] ?? '').trim() : '',
        last: idxLast >= 0 ? String(row[idxLast] ?? '').trim() : '',
        role: idxRole >= 0 ? canonicalizeRoleLabel(row[idxRole]) : '',
        teams,
        teamRaw
      }
    };
  }

  return { exists: false };
}

export function getViewerProfile(): ViewerProfile {
  let viewerEmail = '';
  try {
    const sessionEmail = Session.getActiveUser?.().getEmail?.();
    viewerEmail = normalizeEmail(sessionEmail);
  } catch (_) {
    viewerEmail = '';
  }

  const emptyProfile: ViewerProfile = {
    email: viewerEmail,
    permissions: '',
    first: '',
    last: '',
    isAdmin: false,
    isLoggedIn: !!viewerEmail,
    capabilities: { ...emptyCapabilities }
  };

  if (!viewerEmail) {
    return emptyProfile;
  }

  const sh = getSheetByName(ROLES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return emptyProfile;
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const col = (name: string) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());
  const idxEmail = col(ROLES_COL.email);
  if (idxEmail < 0) throw new Error(`Column "${ROLES_COL.email}" not found on Roles sheet.`);
  const idxPerms = col(ROLES_COL.permissions);
  const idxFirst = col(ROLES_COL.first);
  const idxLast = col(ROLES_COL.last);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (const row of body) {
    const rowEmail = normalizeEmail(row[idxEmail]);
    if (rowEmail !== viewerEmail) continue;
    const permissions = idxPerms >= 0 ? String(row[idxPerms] ?? '').trim() : '';
    const first = idxFirst >= 0 ? String(row[idxFirst] ?? '').trim() : '';
    const last = idxLast >= 0 ? String(row[idxLast] ?? '').trim() : '';
    const { capabilities, isAdmin } = computeCapabilities(permissions);
    return {
      email: viewerEmail,
      permissions,
      first,
      last,
      isAdmin,
      isLoggedIn: true,
      capabilities
    };
  }

  return emptyProfile;
}

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

  const roleFixes: Array<{ row: number; value: string }> = [];

  for (let i = 0; i < body.length; i++) {
    const row = body[i];
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
      role: idxRole >= 0 ? canonicalizeRoleLabel(row[idxRole]) : '',
      spanish: idxSpanish >= 0 ? String(row[idxSpanish] ?? '').trim() : ''
    });
    if (idxRole >= 0) {
      const rawRole = String(row[idxRole] ?? '').trim();
      const canonical = canonicalizeRoleLabel(rawRole);
      if (rawRole && canonical && rawRole !== canonical) {
        roleFixes.push({ row: 2 + i, value: canonical });
      }
    }
  }

  items.sort((a, b) => a.last.localeCompare(b.last) || a.first.localeCompare(b.first));

  if (roleFixes.length && idxRole >= 0) {
    const lock = acquireRolesLock();
    try {
      roleFixes.forEach(fix => {
        sh.getRange(fix.row, idxRole + 1).setValue(fix.value);
      });
    } finally {
      lock.releaseLock();
    }
  }

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

  const lock = acquireRolesLock();
  try {
    const rowNumber = 2 + targetIndex;
    const updates: Array<{ column: number; value: string }> = [];
    if (typeof input.team === 'string') {
      updates.push({ column: idxTeam + 1, value: input.team.trim() });
    }
    if (typeof input.role === 'string') {
      updates.push({ column: idxRole + 1, value: canonicalizeRoleLabel(input.role) });
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
  const lock = acquireRolesLock();
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
      sh.getRange(nextRow, idxRole + 1).setValue(canonicalizeRoleLabel(input.role));
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




