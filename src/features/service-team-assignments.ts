import {
  SERVICE_TEAM_ASSIGNMENTS_SHEET,
  SERVICE_TEAM_ASSIGNMENTS_COL,
  MEMBER_AVAILABILITY_SHEET,
  MEMBER_AVAILABILITY_COL
} from '../constants';
import { listServices, ListServicesOptions, ServiceItem } from './services';

type ServiceTeamAssignmentRow = {
  serviceId: string;
  serviceDate: string;
  serviceType: string;
  teamType: string;
  weeklyTeamName: string;
  roleName: string;
  roleType: string;
  memberEmail: string;
  memberName: string;
  status: string;
  notes: string;
  rowNumber: number;
};

type SaveAssignmentInput = {
  serviceId: string;
  serviceDate?: string;
  serviceType?: string;
  teamType: string;
  weeklyTeamName?: string;
  roleName: string;
  roleType?: string;
  memberEmail?: string;
  memberName?: string;
  status?: string;
  notes?: string;
};

const norm = (value: unknown): string => String(value ?? '').trim();
const normLower = (value: unknown): string => norm(value).toLowerCase();
const assignmentKey = (serviceId: unknown, teamType: unknown, roleName: unknown) =>
  `${normLower(serviceId)}::${normLower(teamType)}::${normLower(roleName)}`;

function headerIndex(headers: string[], label: string) {
  const idx = headers.findIndex(h => h.trim().toLowerCase() === label.trim().toLowerCase());
  if (idx < 0) throw new Error(`Column "${label}" not found on ${SERVICE_TEAM_ASSIGNMENTS_SHEET}.`);
  return idx;
}

function headerIndexOptional(headers: string[], label: string) {
  return headers.findIndex(h => h.trim().toLowerCase() === label.trim().toLowerCase());
}

function toISODate(value: any): string {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const raw = String(value ?? '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  return '';
}

function sheetDateValue(iso: string): Date | string {
  if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso || '';
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, (m || 1) - 1, d || 1);
}

function ensureAssignmentSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SERVICE_TEAM_ASSIGNMENTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SERVICE_TEAM_ASSIGNMENTS_SHEET);
    sh.getRange(1, 1, 1, 11).setValues([[
      SERVICE_TEAM_ASSIGNMENTS_COL.serviceId,
      SERVICE_TEAM_ASSIGNMENTS_COL.serviceDate,
      SERVICE_TEAM_ASSIGNMENTS_COL.serviceType,
      SERVICE_TEAM_ASSIGNMENTS_COL.teamType,
      SERVICE_TEAM_ASSIGNMENTS_COL.weeklyTeamName,
      SERVICE_TEAM_ASSIGNMENTS_COL.roleName,
      SERVICE_TEAM_ASSIGNMENTS_COL.roleType,
      SERVICE_TEAM_ASSIGNMENTS_COL.memberEmail,
      SERVICE_TEAM_ASSIGNMENTS_COL.memberName,
      SERVICE_TEAM_ASSIGNMENTS_COL.status,
      SERVICE_TEAM_ASSIGNMENTS_COL.notes
    ]]);
  }
  return sh;
}

function readAssignmentRows(): ServiceTeamAssignmentRow[] {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SERVICE_TEAM_ASSIGNMENTS_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = sh
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(v => String(v ?? '').trim());

  const idxServiceId = headerIndex(headers, SERVICE_TEAM_ASSIGNMENTS_COL.serviceId);
  const idxServiceDate = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.serviceDate);
  const idxServiceType = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.serviceType);
  const idxTeamType = headerIndex(headers, SERVICE_TEAM_ASSIGNMENTS_COL.teamType);
  const idxWeeklyTeam = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.weeklyTeamName);
  const idxRoleName = headerIndex(headers, SERVICE_TEAM_ASSIGNMENTS_COL.roleName);
  const idxRoleType = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.roleType);
  const idxMemberEmail = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.memberEmail);
  const idxMemberName = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.memberName);
  const idxStatus = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.status);
  const idxNotes = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.notes);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows: ServiceTeamAssignmentRow[] = [];
  body.forEach((row, i) => {
    const serviceId = norm(row[idxServiceId]);
    const teamType = norm(row[idxTeamType]);
    const roleName = norm(row[idxRoleName]);
    if (!serviceId || !teamType || !roleName) return;
    rows.push({
      serviceId,
      serviceDate: idxServiceDate >= 0 ? toISODate(row[idxServiceDate]) : '',
      serviceType: idxServiceType >= 0 ? norm(row[idxServiceType]) : '',
      teamType,
      weeklyTeamName: idxWeeklyTeam >= 0 ? norm(row[idxWeeklyTeam]) : '',
      roleName,
      roleType: idxRoleType >= 0 ? norm(row[idxRoleType]) : '',
      memberEmail: idxMemberEmail >= 0 ? norm(row[idxMemberEmail]) : '',
      memberName: idxMemberName >= 0 ? norm(row[idxMemberName]) : '',
      status: idxStatus >= 0 ? norm(row[idxStatus]) : '',
      notes: idxNotes >= 0 ? norm(row[idxNotes]) : '',
      rowNumber: i + 2
    });
  });
  return rows;
}

function readUnavailableByService(serviceIds: string[]) {
  const idSet = new Set(serviceIds.filter(Boolean));
  if (!idSet.size) return {} as Record<string, Set<string>>;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MEMBER_AVAILABILITY_SHEET);
  if (!sh) return {};
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return {};
  const headers = sh
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(v => String(v ?? '').trim());
  const idxEmail = headerIndex(headers, MEMBER_AVAILABILITY_COL.email);
  const idxServiceId = headerIndex(headers, MEMBER_AVAILABILITY_COL.serviceId);
  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const map: Record<string, Set<string>> = {};
  body.forEach(row => {
    const serviceId = norm(row[idxServiceId]);
    if (!serviceId || !idSet.has(serviceId)) return;
    const email = normLower(row[idxEmail]);
    if (!email) return;
    if (!map[serviceId]) map[serviceId] = new Set();
    map[serviceId].add(email);
  });
  return map;
}

const ordinalSuffix = (value: number) => {
  const j = value % 10;
  const k = value % 100;
  if (j === 1 && k !== 11) return `${value}st`;
  if (j === 2 && k !== 12) return `${value}nd`;
  if (j === 3 && k !== 13) return `${value}rd`;
  return `${value}th`;
};

function ordinalLabelForDate(iso: string, weekday = 'Sunday') {
  const parts = iso.split('-').map(Number);
  if (parts.length !== 3 || parts.some(n => isNaN(n))) return '';
  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  if (isNaN(date.getTime())) return '';
  const day = date.getDate();
  const ordinal = ordinalSuffix(Math.floor((day - 1) / 7) + 1);
  return `${ordinal} ${weekday}`;
}

function formatDateLabel(iso: string, format: string) {
  if (!iso) return '';
  const parts = iso.split('-').map(Number);
  if (parts.length !== 3 || parts.some(n => isNaN(n))) return iso;
  const date = new Date(parts[0], parts[1] - 1, parts[2]);
  if (isNaN(date.getTime())) return iso;
  const tz = (Session.getScriptTimeZone && Session.getScriptTimeZone()) || 'Etc/UTC';
  return Utilities.formatDate(date, tz, format);
}

type TeamScheduleSnapshot = {
  services: Array<{
    id: string;
    date: string;
    time: string;
    type: string;
    label: string;
    shortLabel: string;
    ordinalLabel: string;
    weekday: string;
  }>;
  assignments: Array<Omit<ServiceTeamAssignmentRow, 'rowNumber'>>;
  unavailable: Record<string, string[]>;
};

export function getTeamScheduleSnapshot(input?: { limit?: number } & ListServicesOptions): TeamScheduleSnapshot {
  const limitRaw = Number(input?.limit);
  const limit = Number.isFinite(limitRaw) ? Math.min(12, Math.max(1, Math.floor(limitRaw))) : 6;
  const serviceOptions: ListServicesOptions = {
    includePast: false,
    sort: 'asc',
    limit
  };
  const serviceResult = listServices(serviceOptions);
  const services = Array.isArray(serviceResult?.items) ? serviceResult.items : [];
  const normalizedServices = services
    .filter((svc: ServiceItem) => svc?.id)
    .map(svc => {
      const weekday = formatDateLabel(svc.date, 'EEEE') || 'Sunday';
      return {
        id: svc.id,
        date: svc.date,
        time: svc.time,
        type: svc.type,
        label: formatDateLabel(svc.date, 'EEE, MMM d') || svc.date,
        shortLabel: formatDateLabel(svc.date, 'MMM d') || svc.date,
        ordinalLabel: ordinalLabelForDate(svc.date, weekday),
        weekday
      };
    });
  const serviceIds = normalizedServices.map(svc => svc.id);
  const unavailableMap = readUnavailableByService(serviceIds);
  const assignments = readAssignmentRows().filter(row => serviceIds.includes(row.serviceId));
  const rowsToDelete: number[] = [];
  const cleanedAssignments: ServiceTeamAssignmentRow[] = [];
  assignments.forEach(row => {
    const email = normLower(row.memberEmail);
    if (email && unavailableMap[row.serviceId]?.has(email)) {
      rowsToDelete.push(row.rowNumber);
      return;
    }
    cleanedAssignments.push(row);
  });
  if (rowsToDelete.length) {
    const sh = ensureAssignmentSheet();
    rowsToDelete.sort((a, b) => b - a).forEach(row => sh.deleteRow(row));
  }
  const unavailable = Object.fromEntries(
    Object.entries(unavailableMap).map(([serviceId, set]) => [serviceId, Array.from(set.values())])
  );
  return {
    services: normalizedServices,
    assignments: cleanedAssignments.map(({ rowNumber, ...rest }) => rest),
    unavailable
  };
}

type ServiceTeamAssignmentGroup = {
  teamType: string;
  weeklyTeamName: string;
  roles: Array<{
    roleName: string;
    roleType: string;
    memberEmail: string;
    memberName: string;
    status: string;
    notes: string;
  }>;
};

export function getServiceTeamAssignments(input: { serviceId?: string }) {
  const serviceId = norm(input?.serviceId);
  if (!serviceId) throw new Error('Service ID is required.');
  const rows = readAssignmentRows().filter(row => row.serviceId === serviceId);
  const groups = new Map<string, ServiceTeamAssignmentGroup>();
  rows.forEach(row => {
    const key = normLower(row.teamType) || 'team';
    if (!groups.has(key)) {
      groups.set(key, {
        teamType: row.teamType || 'Team',
        weeklyTeamName: row.weeklyTeamName || '',
        roles: []
      });
    }
    const group = groups.get(key)!;
    if (!group.weeklyTeamName && row.weeklyTeamName) group.weeklyTeamName = row.weeklyTeamName;
    group.roles.push({
      roleName: row.roleName || row.roleType || 'Role',
      roleType: row.roleType || '',
      memberEmail: row.memberEmail || '',
      memberName: row.memberName || '',
      status: row.status || '',
      notes: row.notes || ''
    });
  });
  const teams = Array.from(groups.values())
    .map(group => ({
      ...group,
      roles: group.roles.sort((a, b) => a.roleName.localeCompare(b.roleName))
    }))
    .sort((a, b) => groupLabel(a).localeCompare(groupLabel(b)));
  return { serviceId, teams };
}

function groupLabel(group: ServiceTeamAssignmentGroup): string {
  return (group.teamType || '').toLowerCase();
}

export function saveServiceTeamAssignments(payload: { assignments?: SaveAssignmentInput[] }) {
  const entries = Array.isArray(payload?.assignments) ? payload.assignments : [];
  if (!entries.length) return { updated: 0 };
  const normalized = new Map<string, SaveAssignmentInput & { key: string }>();
  entries.forEach(entry => {
    const serviceId = norm(entry.serviceId);
    const teamType = norm(entry.teamType);
    const roleName = norm(entry.roleName);
    if (!serviceId || !teamType || !roleName) return;
    const key = assignmentKey(serviceId, teamType, roleName);
    normalized.set(key, {
      ...entry,
      serviceId,
      teamType,
      roleName,
      key
    });
  });
  if (!normalized.size) return { updated: 0 };

  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sh = ensureAssignmentSheet();
    const lastCol = Math.max(
      sh.getLastColumn(),
      Object.keys(SERVICE_TEAM_ASSIGNMENTS_COL).length
    );
    const headers = sh
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map(v => String(v ?? '').trim());

    const idxServiceId = headerIndex(headers, SERVICE_TEAM_ASSIGNMENTS_COL.serviceId);
    const idxServiceDate = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.serviceDate);
    const idxServiceType = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.serviceType);
    const idxTeamType = headerIndex(headers, SERVICE_TEAM_ASSIGNMENTS_COL.teamType);
    const idxWeeklyTeam = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.weeklyTeamName);
    const idxRoleName = headerIndex(headers, SERVICE_TEAM_ASSIGNMENTS_COL.roleName);
    const idxRoleType = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.roleType);
    const idxMemberEmail = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.memberEmail);
    const idxMemberName = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.memberName);
    const idxStatus = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.status);
    const idxNotes = headerIndexOptional(headers, SERVICE_TEAM_ASSIGNMENTS_COL.notes);

    const existing = readAssignmentRows();
    const existingMap = new Map(existing.map(row => [assignmentKey(row.serviceId, row.teamType, row.roleName), row]));

    const rowsToDelete: number[] = [];
    const rowUpdates: Array<{ rowNumber: number; entry: SaveAssignmentInput & { key: string } }> = [];
    const rowsToInsert: string[][] = [];

    normalized.forEach(entry => {
      const key = entry.key;
      const memberEmail = norm(entry.memberEmail);
      const existingRow = existingMap.get(key);
      if (!memberEmail) {
        if (existingRow) {
          rowsToDelete.push(existingRow.rowNumber);
        }
        return;
      }
      if (existingRow) {
        rowUpdates.push({ rowNumber: existingRow.rowNumber, entry });
      } else {
        const colCount = Math.max(headers.length, sh.getLastColumn());
        const rowValues = Array.from({ length: colCount }, () => '');
        rowValues[idxServiceId] = entry.serviceId;
        if (idxServiceDate >= 0) rowValues[idxServiceDate] = sheetDateValue(norm(entry.serviceDate));
        if (idxServiceType >= 0) rowValues[idxServiceType] = norm(entry.serviceType);
        rowValues[idxTeamType] = entry.teamType;
        if (idxWeeklyTeam >= 0) rowValues[idxWeeklyTeam] = norm(entry.weeklyTeamName);
        rowValues[idxRoleName] = entry.roleName;
        if (idxRoleType >= 0) rowValues[idxRoleType] = norm(entry.roleType || entry.roleName);
        if (idxMemberEmail >= 0) rowValues[idxMemberEmail] = memberEmail;
        if (idxMemberName >= 0) rowValues[idxMemberName] = norm(entry.memberName);
        if (idxStatus >= 0) rowValues[idxStatus] = norm(entry.status) || 'Assigned';
        if (idxNotes >= 0) rowValues[idxNotes] = norm(entry.notes);
        rowsToInsert.push(rowValues);
      }
    });

    rowsToDelete.sort((a, b) => b - a).forEach(row => sh.deleteRow(row));
    rowUpdates.forEach(({ rowNumber, entry }) => {
      if (idxServiceId >= 0) sh.getRange(rowNumber, idxServiceId + 1).setValue(entry.serviceId);
      if (idxServiceDate >= 0) sh.getRange(rowNumber, idxServiceDate + 1).setValue(sheetDateValue(norm(entry.serviceDate)));
      if (idxServiceType >= 0) sh.getRange(rowNumber, idxServiceType + 1).setValue(norm(entry.serviceType));
      if (idxTeamType >= 0) sh.getRange(rowNumber, idxTeamType + 1).setValue(entry.teamType);
      if (idxWeeklyTeam >= 0) sh.getRange(rowNumber, idxWeeklyTeam + 1).setValue(norm(entry.weeklyTeamName));
      if (idxRoleName >= 0) sh.getRange(rowNumber, idxRoleName + 1).setValue(entry.roleName);
      if (idxRoleType >= 0) sh.getRange(rowNumber, idxRoleType + 1).setValue(norm(entry.roleType || entry.roleName));
      if (idxMemberEmail >= 0) sh.getRange(rowNumber, idxMemberEmail + 1).setValue(norm(entry.memberEmail));
      if (idxMemberName >= 0) sh.getRange(rowNumber, idxMemberName + 1).setValue(norm(entry.memberName));
      if (idxStatus >= 0) sh.getRange(rowNumber, idxStatus + 1).setValue(norm(entry.status) || 'Assigned');
      if (idxNotes >= 0) sh.getRange(rowNumber, idxNotes + 1).setValue(norm(entry.notes));
    });
    if (rowsToInsert.length) {
      const startRow = Math.max(sh.getLastRow(), 1) + 1;
      sh.getRange(startRow, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    }
    return { updated: rowsToDelete.length + rowUpdates.length + rowsToInsert.length };
  } finally {
    lock.releaseLock();
  }
}
