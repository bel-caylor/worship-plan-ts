import {
  VOLUNTEER_REQUESTS_COL,
  VOLUNTEER_REQUESTS_SHEET
} from '../constants';
import { getAvailabilityIndex } from './member-availability';
import { getViewerProfile, findRoleRecordByEmail } from './roles';
import { getTeamScheduleSnapshot } from './service-team-assignments';

type VolunteerRequestRow = {
  serviceId: string;
  teamType: string;
  roleName: string;
  memberEmail: string;
  memberName: string;
  status: string;
  requestedAt: string;
  notes: string;
  rowNumber: number;
};

const norm = (value: unknown) => String(value ?? '').trim();
const normLower = (value: unknown) => norm(value).toLowerCase();
const slotKey = (serviceId: unknown, teamType: unknown, roleName: unknown) =>
  `${normLower(serviceId)}::${normLower(teamType)}::${normLower(roleName)}`;
const canonicalRole = (value: unknown) => {
  const role = norm(value);
  if (!role) return '';
  return /^vocals?$/i.test(role) ? 'Vocal' : role;
};
const isoNow = () => new Date().toISOString();

function headerIndex(headers: string[], label: string) {
  const idx = headers.findIndex(h => h.trim().toLowerCase() === label.trim().toLowerCase());
  if (idx < 0) throw new Error(`Column "${label}" not found on ${VOLUNTEER_REQUESTS_SHEET}.`);
  return idx;
}

function headerIndexOptional(headers: string[], label: string) {
  return headers.findIndex(h => h.trim().toLowerCase() === label.trim().toLowerCase());
}

function ensureVolunteerRequestsSheet(): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(VOLUNTEER_REQUESTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(VOLUNTEER_REQUESTS_SHEET);
    sh.getRange(1, 1, 1, 8).setValues([[
      VOLUNTEER_REQUESTS_COL.serviceId,
      VOLUNTEER_REQUESTS_COL.teamType,
      VOLUNTEER_REQUESTS_COL.roleName,
      VOLUNTEER_REQUESTS_COL.memberEmail,
      VOLUNTEER_REQUESTS_COL.memberName,
      VOLUNTEER_REQUESTS_COL.status,
      VOLUNTEER_REQUESTS_COL.requestedAt,
      VOLUNTEER_REQUESTS_COL.notes
    ]]);
  }
  return sh;
}

function readVolunteerRequestRows(): VolunteerRequestRow[] {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(VOLUNTEER_REQUESTS_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
  const idxServiceId = headerIndex(headers, VOLUNTEER_REQUESTS_COL.serviceId);
  const idxTeamType = headerIndex(headers, VOLUNTEER_REQUESTS_COL.teamType);
  const idxRoleName = headerIndex(headers, VOLUNTEER_REQUESTS_COL.roleName);
  const idxMemberEmail = headerIndex(headers, VOLUNTEER_REQUESTS_COL.memberEmail);
  const idxMemberName = headerIndexOptional(headers, VOLUNTEER_REQUESTS_COL.memberName);
  const idxStatus = headerIndexOptional(headers, VOLUNTEER_REQUESTS_COL.status);
  const idxRequestedAt = headerIndexOptional(headers, VOLUNTEER_REQUESTS_COL.requestedAt);
  const idxNotes = headerIndexOptional(headers, VOLUNTEER_REQUESTS_COL.notes);

  const body = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return body
    .map((row, index) => ({
      serviceId: norm(row[idxServiceId]),
      teamType: norm(row[idxTeamType]),
      roleName: canonicalRole(row[idxRoleName]),
      memberEmail: normLower(row[idxMemberEmail]),
      memberName: idxMemberName >= 0 ? norm(row[idxMemberName]) : '',
      status: idxStatus >= 0 ? norm(row[idxStatus]) : '',
      requestedAt: idxRequestedAt >= 0 ? norm(row[idxRequestedAt]) : '',
      notes: idxNotes >= 0 ? norm(row[idxNotes]) : '',
      rowNumber: index + 2
    }))
    .filter(row => row.serviceId && row.teamType && row.roleName && row.memberEmail);
}

function viewerCanVolunteerForSlot(
  viewer: ReturnType<typeof getViewerProfile>,
  teamType: string,
  roleName: string
) {
  if (!viewer?.isLoggedIn || !viewer?.capabilities?.canVolunteer) return false;
  const viewerRole = normLower(canonicalRole(viewer.role));
  const slotRole = normLower(canonicalRole(roleName));
  if (!viewerRole || viewerRole !== slotRole) return false;
  const teamKey = normLower(teamType);
  const teams = Array.isArray(viewer.teams) ? viewer.teams.map(normLower).filter(Boolean) : [];
  return !!teamKey && teams.includes(teamKey);
}

function ensureViewerEligibility(teamType: string, roleName: string) {
  const viewer = getViewerProfile();
  if (!viewer?.isLoggedIn || !viewer.email) {
    throw new Error('Please sign in with Google before volunteering.');
  }
  const roleRecord = findRoleRecordByEmail(viewer.email);
  if (!roleRecord) {
    throw new Error('Your email is not registered in the Roles sheet.');
  }
  if (!viewerCanVolunteerForSlot(viewer, teamType, roleName)) {
    throw new Error('This slot is not available for your current team/role.');
  }
  const unavailable = getAvailabilityIndex().byEmail?.[normLower(viewer.email)] || [];
  return {
    viewer,
    unavailableServiceIds: new Set((Array.isArray(unavailable) ? unavailable : []).map(id => norm(id)).filter(Boolean))
  };
}

function ensureSlotOpen(serviceId: string, teamType: string, roleName: string) {
  const snapshot = getTeamScheduleSnapshot({ limit: 0 });
  const match = (Array.isArray(snapshot?.assignments) ? snapshot.assignments : []).find(entry =>
    slotKey(entry?.serviceId, entry?.teamType, entry?.roleName) === slotKey(serviceId, teamType, roleName)
  );
  if (!match) {
    throw new Error('That schedule slot could not be found.');
  }
  if (norm(match.memberEmail)) {
    throw new Error('That slot has already been assigned.');
  }
  return match;
}

export function getVolunteerRequestsSnapshot(input?: { serviceIds?: string[] }) {
  const serviceIds = Array.isArray(input?.serviceIds)
    ? input.serviceIds.map(id => norm(id)).filter(Boolean)
    : [];
  const serviceIdSet = serviceIds.length ? new Set(serviceIds) : null;
  const viewer = getViewerProfile();
  const viewerEmail = normLower(viewer?.email);
  const items = readVolunteerRequestRows()
    .filter(row => row.status.toLowerCase() !== 'withdrawn')
    .filter(row => !serviceIdSet || serviceIdSet.has(row.serviceId))
    .map(row => ({
      serviceId: row.serviceId,
      teamType: row.teamType,
      roleName: row.roleName,
      memberEmail: row.memberEmail,
      memberName: row.memberName,
      status: row.status || 'Requested',
      requestedAt: row.requestedAt,
      isViewer: !!viewerEmail && row.memberEmail === viewerEmail
    }));
  return { items };
}

export function setViewerVolunteerRequest(input?: {
  serviceId?: string;
  teamType?: string;
  roleName?: string;
  requested?: boolean;
}) {
  const serviceId = norm(input?.serviceId);
  const teamType = norm(input?.teamType);
  const roleName = canonicalRole(input?.roleName);
  const requested = input?.requested !== false;
  if (!serviceId || !teamType || !roleName) {
    throw new Error('Service, team, and role are required.');
  }

  const { viewer, unavailableServiceIds } = ensureViewerEligibility(teamType, roleName);
  if (unavailableServiceIds.has(serviceId)) {
    throw new Error('You marked yourself unavailable for this service.');
  }
  ensureSlotOpen(serviceId, teamType, roleName);

  const email = normLower(viewer.email);
  const name = [norm(viewer.first), norm(viewer.last)].filter(Boolean).join(' ') || email;
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sh = ensureVolunteerRequestsSheet();
    const lastCol = Math.max(sh.getLastColumn(), Object.keys(VOLUNTEER_REQUESTS_COL).length);
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v ?? '').trim());
    const idxServiceId = headerIndex(headers, VOLUNTEER_REQUESTS_COL.serviceId);
    const idxTeamType = headerIndex(headers, VOLUNTEER_REQUESTS_COL.teamType);
    const idxRoleName = headerIndex(headers, VOLUNTEER_REQUESTS_COL.roleName);
    const idxMemberEmail = headerIndex(headers, VOLUNTEER_REQUESTS_COL.memberEmail);
    const idxMemberName = headerIndexOptional(headers, VOLUNTEER_REQUESTS_COL.memberName);
    const idxStatus = headerIndexOptional(headers, VOLUNTEER_REQUESTS_COL.status);
    const idxRequestedAt = headerIndexOptional(headers, VOLUNTEER_REQUESTS_COL.requestedAt);
    const idxNotes = headerIndexOptional(headers, VOLUNTEER_REQUESTS_COL.notes);

    const existing = readVolunteerRequestRows().find(row =>
      row.memberEmail === email && slotKey(row.serviceId, row.teamType, row.roleName) === slotKey(serviceId, teamType, roleName)
    );

    if (!requested) {
      if (existing) {
        sh.deleteRow(existing.rowNumber);
      }
      return { requested: false };
    }

    if (existing) {
      if (idxStatus >= 0) sh.getRange(existing.rowNumber, idxStatus + 1).setValue('Requested');
      if (idxRequestedAt >= 0) sh.getRange(existing.rowNumber, idxRequestedAt + 1).setValue(isoNow());
      if (idxMemberName >= 0) sh.getRange(existing.rowNumber, idxMemberName + 1).setValue(name);
      return { requested: true };
    }

    const rowValues = Array.from({ length: lastCol }, () => '');
    rowValues[idxServiceId] = serviceId;
    rowValues[idxTeamType] = teamType;
    rowValues[idxRoleName] = roleName;
    rowValues[idxMemberEmail] = email;
    if (idxMemberName >= 0) rowValues[idxMemberName] = name;
    if (idxStatus >= 0) rowValues[idxStatus] = 'Requested';
    if (idxRequestedAt >= 0) rowValues[idxRequestedAt] = isoNow();
    if (idxNotes >= 0) rowValues[idxNotes] = '';

    const startRow = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(startRow, 1, 1, rowValues.length).setValues([rowValues]);
    return { requested: true };
  } finally {
    lock.releaseLock();
  }
}
