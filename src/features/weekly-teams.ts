import {
  WEEKLY_TEAMS_SHEET,
  WEEKLY_TEAMS_COL,
  WEEKLY_TEAM_ROLES_SHEET,
  WEEKLY_TEAM_ROLES_COL,
  WEEKLY_TEAM_ROLE_DEFAULTS_SHEET,
  WEEKLY_TEAM_ROLE_DEFAULTS_COL
} from '../constants';
import { getSheetByName } from '../util/sheets';

type WeeklyTeamRole = {
  roleType: string;
  roleName: string;
  memberEmail: string;
  memberName: string;
};

type WeeklyTeamRecord = {
  team: string;
  teamName: string;
  description: string;
  roles: WeeklyTeamRole[];
};

type WeeklyTeamSheetRow = WeeklyTeamRecord & {
  key: string;
  rowNumber: number;
};

type WeeklyTeamRoleSheetRow = WeeklyTeamRole & {
  team: string;
  teamName: string;
  key: string;
  rowNumber: number;
};

type WeeklyTeamRoleDefault = {
  team: string;
  roleName: string;
  order: number;
  rowNumber: number;
};

type SaveWeeklyTeamInput = {
  team: string;
  teamName: string;
  description?: string;
  roles?: Array<{
    roleName?: string;
    roleType?: string;
    memberEmail?: string;
    memberName?: string;
  }>;
  original?: {
    team?: string;
    teamName?: string;
  };
};

type CreateWeeklyTeamInput = {
  team: string;
  teamName: string;
  description?: string;
};

const norm = (value: unknown): string => String(value ?? '').trim();
const normKey = (team: unknown, teamName: unknown): string =>
  `${norm(team).toLowerCase()}::${norm(teamName).toLowerCase()}`;

function acquireDocumentLock(maxAttempts = 4, waitMs = 5000) {
  const baseDelay = 250;
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

function getSheetOrNull(name: string): GoogleAppsScript.Spreadsheet.Sheet | null {
  try {
    return getSheetByName(name);
  } catch (_) {
    return null;
  }
}

function headerIndex(headers: string[], label: string) {
  const idx = headers.findIndex(h => h.trim().toLowerCase() === label.trim().toLowerCase());
  if (idx < 0) throw new Error(`Column "${label}" not found.`);
  return idx;
}

function headerIndexOptional(headers: string[], label: string) {
  return headers.findIndex(h => h.trim().toLowerCase() === label.trim().toLowerCase());
}

function readWeeklyTeamsSheet(): WeeklyTeamSheetRow[] {
  const sh = getSheetOrNull(WEEKLY_TEAMS_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = sh
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(v => String(v ?? '').trim());

  const idxTeam = headerIndex(headers, WEEKLY_TEAMS_COL.team);
  const idxTeamName = headerIndex(headers, WEEKLY_TEAMS_COL.teamName);
  const idxDescription = headerIndexOptional(headers, WEEKLY_TEAMS_COL.description);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows: WeeklyTeamSheetRow[] = [];
  body.forEach((row, i) => {
    const team = norm(row[idxTeam]);
    const teamName = norm(row[idxTeamName]);
    if (!team || !teamName) return;
    rows.push({
      team,
      teamName,
      description: idxDescription >= 0 ? norm(row[idxDescription]) : '',
      roles: [],
      key: normKey(team, teamName),
      rowNumber: 2 + i
    });
  });
  return rows;
}

function ensureDefaultsSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(WEEKLY_TEAM_ROLE_DEFAULTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(WEEKLY_TEAM_ROLE_DEFAULTS_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([[
      WEEKLY_TEAM_ROLE_DEFAULTS_COL.team,
      WEEKLY_TEAM_ROLE_DEFAULTS_COL.roleName,
      WEEKLY_TEAM_ROLE_DEFAULTS_COL.order
    ]]);
  }
  return sh;
}

function readWeeklyTeamRoleDefaults(): WeeklyTeamRoleDefault[] {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(WEEKLY_TEAM_ROLE_DEFAULTS_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = sh
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(v => String(v ?? '').trim());

  const idxTeam = headerIndex(headers, WEEKLY_TEAM_ROLE_DEFAULTS_COL.team);
  const idxRoleName = headerIndex(headers, WEEKLY_TEAM_ROLE_DEFAULTS_COL.roleName);
  const idxOrder = headerIndexOptional(headers, WEEKLY_TEAM_ROLE_DEFAULTS_COL.order);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows: WeeklyTeamRoleDefault[] = [];
  body.forEach((row, i) => {
    const team = norm(row[idxTeam]);
    const roleName = norm(row[idxRoleName]);
    if (!team || !roleName) return;
    const orderValue = idxOrder >= 0 ? Number(row[idxOrder]) : NaN;
    rows.push({
      team,
      roleName,
      order: Number.isFinite(orderValue) ? orderValue : i,
      rowNumber: i + 2
    });
  });
  return rows;
}

function readWeeklyTeamRolesSheet(): WeeklyTeamRoleSheetRow[] {
  const sh = getSheetByName(WEEKLY_TEAM_ROLES_SHEET);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = sh
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(v => String(v ?? '').trim());

  const idxTeam = headerIndex(headers, WEEKLY_TEAM_ROLES_COL.team);
  const idxTeamName = headerIndex(headers, WEEKLY_TEAM_ROLES_COL.teamName);
  const idxRoleName = headerIndex(headers, WEEKLY_TEAM_ROLES_COL.roleName);
  const idxRoleType = headerIndexOptional(headers, WEEKLY_TEAM_ROLES_COL.roleType);
  const idxMemberEmail = headerIndex(headers, WEEKLY_TEAM_ROLES_COL.memberEmail);
  const idxMemberName = headerIndex(headers, WEEKLY_TEAM_ROLES_COL.memberName);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows: WeeklyTeamRoleSheetRow[] = [];
  body.forEach((row, i) => {
    const team = norm(row[idxTeam]);
    const teamName = norm(row[idxTeamName]);
    const roleName = norm(row[idxRoleName]);
    if (!team || !teamName || !roleName) return;
    rows.push({
      team,
      teamName,
      roleName,
      roleType: idxRoleType >= 0 ? norm(row[idxRoleType]) : '',
      memberEmail: norm(row[idxMemberEmail]),
      memberName: norm(row[idxMemberName]),
      key: normKey(team, teamName),
      rowNumber: 2 + i
    });
  });
  return rows;
}

export function listWeeklyTeams() {
  const baseRows = readWeeklyTeamsSheet();
  const roleRows = readWeeklyTeamRolesSheet();
  const defaultRows = readWeeklyTeamRoleDefaults();

  const map = new Map<string, WeeklyTeamRecord>();
  baseRows.forEach(row => {
    map.set(row.key, {
      team: row.team,
      teamName: row.teamName,
      description: row.description,
      roles: []
    });
  });

  roleRows.forEach(row => {
    const existing = map.get(row.key) || {
      team: row.team,
      teamName: row.teamName,
      description: '',
      roles: []
    };
    existing.roles.push({
      roleType: row.roleType,
      roleName: row.roleName,
      memberEmail: row.memberEmail,
      memberName: row.memberName
    });
    map.set(row.key, existing);
  });

  const items = Array.from(map.values()).sort((a, b) => {
    const teamCompare = a.team.localeCompare(b.team);
    if (teamCompare !== 0) return teamCompare;
    return a.teamName.localeCompare(b.teamName);
  });

  const defaultsMap = new Map<string, WeeklyTeamRoleDefault[]>();
  defaultRows.forEach(row => {
    const key = norm(row.team).toLowerCase();
    if (!key) return;
    if (!defaultsMap.has(key)) defaultsMap.set(key, []);
    defaultsMap.get(key)!.push(row);
  });
  const defaultsObj: Record<string, { roleName: string; order: number }[]> = {};
  defaultsMap.forEach((list, key) => {
    list.sort((a, b) => {
      if (a.order !== b.order) return a.order - b.order;
      return a.roleName.localeCompare(b.roleName);
    });
    if (!list.length) return;
    const canonicalTeam = list[0].team;
    defaultsObj[canonicalTeam] = list.map(entry => ({
      roleName: entry.roleName,
      order: entry.order
    }));
  });

  return { items, defaults: defaultsObj };
}

export function saveWeeklyTeamDefaults(input: { team: string; roles: string[] }) {
  const team = norm(input.team);
  if (!team) throw new Error('Team type is required.');
  const roles = Array.isArray(input.roles) ? input.roles.map(role => norm(role)).filter(Boolean) : [];

  const lock = acquireDocumentLock();
  try {
    const sh = ensureDefaultsSheet();
    const lastCol = Math.max(sh.getLastColumn(), 3);
    const headers = sh
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map(v => String(v ?? '').trim());

    const idxTeam = headerIndex(headers, WEEKLY_TEAM_ROLE_DEFAULTS_COL.team);
    const idxRole = headerIndex(headers, WEEKLY_TEAM_ROLE_DEFAULTS_COL.roleName);
    let idxOrder = headerIndexOptional(headers, WEEKLY_TEAM_ROLE_DEFAULTS_COL.order);
    if (idxOrder < 0) {
      sh.insertColumnAfter(lastCol);
      idxOrder = lastCol;
      sh.getRange(1, idxOrder + 1).setValue(WEEKLY_TEAM_ROLE_DEFAULTS_COL.order);
    }

    const existing = readWeeklyTeamRoleDefaults();
    const rowsToDelete = existing
      .filter(entry => norm(entry.team).toLowerCase() === team.toLowerCase())
      .map(entry => entry.rowNumber)
      .sort((a, b) => b - a);
    rowsToDelete.forEach(rowNumber => sh.deleteRow(rowNumber));

    roles.forEach((roleName, order) => {
      const nextRow = Math.max(sh.getLastRow(), 1) + 1;
      sh.getRange(nextRow, idxTeam + 1).setValue(team);
      sh.getRange(nextRow, idxRole + 1).setValue(roleName);
      sh.getRange(nextRow, idxOrder + 1).setValue(order);
    });
  } finally {
    lock.releaseLock();
  }

  return listWeeklyTeams();
}

export function createWeeklyTeam(input: CreateWeeklyTeamInput) {
  const team = norm(input.team);
  const teamName = norm(input.teamName);
  const description = norm(input.description);
  if (!team) throw new Error('Team type is required.');
  if (!teamName) throw new Error('Team name is required.');

  const lock = acquireDocumentLock();
  try {
    const existing = readWeeklyTeamsSheet();
    const key = normKey(team, teamName);
    if (existing.some(row => row.key === key)) {
      throw new Error(`Weekly team "${teamName}" already exists for ${team}.`);
    }
    const sh = getSheetByName(WEEKLY_TEAMS_SHEET);
    const lastCol = sh.getLastColumn();
    if (lastCol < 1) throw new Error('WeeklyTeams sheet has no columns.');
    const headers = sh
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map(v => String(v ?? '').trim());

    const idxTeam = headerIndex(headers, WEEKLY_TEAMS_COL.team);
    const idxTeamName = headerIndex(headers, WEEKLY_TEAMS_COL.teamName);
    const idxDescription = headerIndexOptional(headers, WEEKLY_TEAMS_COL.description);

    const rowNumber = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(rowNumber, idxTeam + 1).setValue(team);
    sh.getRange(rowNumber, idxTeamName + 1).setValue(teamName);
    if (idxDescription >= 0) {
      sh.getRange(rowNumber, idxDescription + 1).setValue(description);
    }
  } finally {
    lock.releaseLock();
  }

  return listWeeklyTeams();
}

export function saveWeeklyTeam(input: SaveWeeklyTeamInput) {
  const team = norm(input.team);
  const teamName = norm(input.teamName);
  if (!team) throw new Error('Team type is required.');
  if (!teamName) throw new Error('Team name is required.');

  const description = norm(input.description);
  const originalTeam = norm(input.original?.team);
  const originalTeamName = norm(input.original?.teamName);
  const effectiveOriginalTeam = originalTeam || team;
  const effectiveOriginalTeamName = originalTeamName || teamName;
  const targetKey = normKey(team, teamName);
  const originalKey = normKey(effectiveOriginalTeam, effectiveOriginalTeamName);

  const lock = acquireDocumentLock();
  try {
    const teamSheet = getSheetByName(WEEKLY_TEAMS_SHEET);
    const roleSheet = getSheetByName(WEEKLY_TEAM_ROLES_SHEET);

    const teamHeaders = teamSheet
      .getRange(1, 1, 1, Math.max(teamSheet.getLastColumn(), 1))
      .getValues()[0]
      .map(v => String(v ?? '').trim());
    const idxTeam = headerIndex(teamHeaders, WEEKLY_TEAMS_COL.team);
    const idxTeamName = headerIndex(teamHeaders, WEEKLY_TEAMS_COL.teamName);
    const idxDescription = headerIndexOptional(teamHeaders, WEEKLY_TEAMS_COL.description);

    const existingTeams = readWeeklyTeamsSheet();
    let targetRow = existingTeams.find(row => row.key === originalKey);
    if (!targetRow) {
      targetRow = existingTeams.find(row => row.key === targetKey);
    }

    if (!targetRow) {
      const rowNumber = Math.max(teamSheet.getLastRow(), 1) + 1;
      teamSheet.getRange(rowNumber, idxTeam + 1).setValue(team);
      teamSheet.getRange(rowNumber, idxTeamName + 1).setValue(teamName);
      if (idxDescription >= 0) {
        teamSheet.getRange(rowNumber, idxDescription + 1).setValue(description);
      }
    } else {
      teamSheet.getRange(targetRow.rowNumber, idxTeam + 1).setValue(team);
      teamSheet.getRange(targetRow.rowNumber, idxTeamName + 1).setValue(teamName);
      if (idxDescription >= 0) {
        teamSheet.getRange(targetRow.rowNumber, idxDescription + 1).setValue(description);
      }
    }

    const roleHeaders = roleSheet
      .getRange(1, 1, 1, Math.max(roleSheet.getLastColumn(), 1))
      .getValues()[0]
      .map(v => String(v ?? '').trim());
    const idxRoleTeam = headerIndex(roleHeaders, WEEKLY_TEAM_ROLES_COL.team);
    const idxRoleTeamName = headerIndex(roleHeaders, WEEKLY_TEAM_ROLES_COL.teamName);
    const idxRoleName = headerIndex(roleHeaders, WEEKLY_TEAM_ROLES_COL.roleName);
    const idxRoleType = headerIndexOptional(roleHeaders, WEEKLY_TEAM_ROLES_COL.roleType);
    const idxRoleMemberEmail = headerIndex(roleHeaders, WEEKLY_TEAM_ROLES_COL.memberEmail);
    const idxRoleMemberName = headerIndex(roleHeaders, WEEKLY_TEAM_ROLES_COL.memberName);

    const existingRoles = readWeeklyTeamRolesSheet();
    const rowsToDelete = existingRoles
      .filter(row => row.key === originalKey || row.key === targetKey)
      .map(row => row.rowNumber)
      .sort((a, b) => b - a);
    rowsToDelete.forEach(rowNumber => roleSheet.deleteRow(rowNumber));

    const roleEntries = Array.isArray(input.roles) ? input.roles : [];
    const roleColumnCount = Math.max(roleHeaders.length, roleSheet.getLastColumn());
    const rowsToInsert: string[][] = [];
    roleEntries.forEach(entry => {
      const roleName = norm(entry.roleName);
      if (!roleName) return;
      const roleType = norm(entry.roleType) || roleName;
      const memberEmail = norm(entry.memberEmail);
      const memberName = norm(entry.memberName);
      const row: string[] = Array.from({ length: roleColumnCount }, () => '');
      row[idxRoleTeam] = team;
      row[idxRoleTeamName] = teamName;
      row[idxRoleName] = roleName;
      if (idxRoleType >= 0) row[idxRoleType] = roleType;
      row[idxRoleMemberEmail] = memberEmail;
      row[idxRoleMemberName] = memberName;
      rowsToInsert.push(row);
    });
    if (rowsToInsert.length) {
      const startRow = Math.max(roleSheet.getLastRow(), 1) + 1;
      roleSheet.getRange(startRow, 1, rowsToInsert.length, roleColumnCount).setValues(rowsToInsert);
    }
  } finally {
    lock.releaseLock();
  }

  return listWeeklyTeams();
}
