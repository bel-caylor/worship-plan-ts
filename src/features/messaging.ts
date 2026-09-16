import { getViewerProfile } from './roles';
import { getServiceTeamAssignments } from './service-team-assignments';

type SendAvailabilityEmailInput = {
  recipients: string[];
  subject?: string;
  body?: string;
};

const normalizeEmail = (value: unknown) => String(value ?? '').trim().toLowerCase();

export function sendAvailabilityEmail(input: SendAvailabilityEmailInput) {
  const profile = getViewerProfile();
  if (!profile?.isAdmin) {
    throw new Error('Only admins can send availability emails.');
  }

  const recipients = Array.isArray(input?.recipients)
    ? input!.recipients.map(normalizeEmail).filter(Boolean)
    : [];
  const uniqueRecipients = Array.from(new Set(recipients));
  if (!uniqueRecipients.length) {
    throw new Error('Select at least one recipient.');
  }

  const subject = String(input?.subject ?? '').trim() || 'Please update your availability';
  const body = String(input?.body ?? '').trim();
  if (!body) {
    throw new Error('Email body is required.');
  }

  const viewerEmail = normalizeEmail(profile?.email);
  const primaryRecipient = viewerEmail || uniqueRecipients[0];
  const bccList = viewerEmail
    ? uniqueRecipients.join(', ')
    : uniqueRecipients.slice(1).join(', ');

  MailApp.sendEmail({
    to: primaryRecipient,
    bcc: bccList || undefined,
    subject,
    body,
    name: profile?.first ? `${profile.first} ${profile.last || ''}`.trim() : 'Worship Planner',
    replyTo: viewerEmail || undefined
  });

  return {
    sent: uniqueRecipients.length,
    subject
  };
}

type SendServiceTeamEmailInput = {
  serviceId?: string;
  subject?: string;
  body?: string;
  recipients?: string[];
  htmlBody?: string;
  requestId?: string;
};

const TEAM_EMAIL_HISTORY_PROPERTY = 'teamEmailSendHistory:v1';
const TEAM_EMAIL_HISTORY_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;
const TEAM_EMAIL_HISTORY_LIMIT = 100;

type TeamEmailSendResult = { sent: number; subject: string; serviceId: string; requestId: string };

function readTeamEmailHistory(): Record<string, TeamEmailSendResult & { completedAt: number }> {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(TEAM_EMAIL_HISTORY_PROPERTY);
    const parsed = raw ? JSON.parse(raw) : null;
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch (_) {
    return {};
  }
}

function writeTeamEmailHistory(history: Record<string, TeamEmailSendResult & { completedAt: number }>) {
  const cutoff = Date.now() - TEAM_EMAIL_HISTORY_MAX_AGE_MS;
  const entries = Object.entries(history)
    .filter(([, value]) => Number(value?.completedAt || 0) >= cutoff)
    .sort(([, a], [, b]) => Number(b.completedAt || 0) - Number(a.completedAt || 0))
    .slice(0, TEAM_EMAIL_HISTORY_LIMIT);
  PropertiesService.getScriptProperties().setProperty(TEAM_EMAIL_HISTORY_PROPERTY, JSON.stringify(Object.fromEntries(entries)));
}

export function sendServiceTeamEmail(input: SendServiceTeamEmailInput) {
  const profile = getViewerProfile();
  const canEmail = Boolean(profile?.capabilities?.canEditPlan);
  if (!canEmail) {
    throw new Error('Only planners can email the serving team.');
  }

  const serviceId = String(input?.serviceId || '').trim();
  if (!serviceId) throw new Error('Service ID is required.');
  const requestId = String(input?.requestId || '').trim();
  if (!requestId || requestId.length > 160) throw new Error('A valid email request ID is required.');

  // The browser retains this ID if the proxy loses the response. A repeat
  // request can then return this completed result without sending again.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const history = readTeamEmailHistory();
    const prior = history[requestId];
    if (prior) return prior;

  const subject = String(input?.subject || '').trim() || 'Team assignments';
  const body = String(input?.body || '').trim();
  const htmlBody = String(input?.htmlBody || '').trim();
  if (!body) throw new Error('Email body is required.');

  const recipientMap = new Map<string, string>();
  const addRecipient = (email: string | null | undefined) => {
    const normalized = normalizeEmail(email);
    if (!normalized) return;
    if (!recipientMap.has(normalized)) {
      const original = String(email || '').trim() || normalized;
      recipientMap.set(normalized, original);
    }
  };

  if (Array.isArray(input?.recipients)) {
    input!.recipients.forEach(email => addRecipient(email));
  }

  if (!recipientMap.size) {
    try {
      const assignments = getServiceTeamAssignments({ serviceId });
      const teams = Array.isArray(assignments?.teams) ? assignments.teams : [];
      teams.forEach(team => {
        (Array.isArray(team?.roles) ? team.roles : []).forEach(role => addRecipient(role?.memberEmail));
      });
    } catch (err) {
      try {
        Logger.log(`team email lookup failed for ${serviceId}: ${err}`);
      } catch (_) { /* ignore */ }
    }
  }

  if (!recipientMap.size) {
    throw new Error('Add at least one team member with an email before sending.');
  }

  const uniqueRecipients = Array.from(recipientMap.values());
  const viewerEmail = normalizeEmail(profile?.email);
  const senderName = profile?.first ? `${profile.first} ${profile.last || ''}`.trim() : 'Worship Planner';

  const toAddress = viewerEmail || uniqueRecipients[0];
  const normalizedTo = normalizeEmail(toAddress);
  const remaining = uniqueRecipients.filter(email => normalizeEmail(email) !== normalizedTo);
  const bccList = viewerEmail
    ? uniqueRecipients.filter(email => normalizeEmail(email) !== viewerEmail)
    : remaining;

    MailApp.sendEmail({
    to: toAddress,
    bcc: bccList.length ? bccList.join(', ') : undefined,
    subject,
    body,
    htmlBody: htmlBody || undefined,
    name: senderName || undefined,
    replyTo: viewerEmail || undefined
  });

    const result: TeamEmailSendResult & { completedAt: number } = {
      sent: uniqueRecipients.length,
      subject,
      serviceId,
      requestId,
      completedAt: Date.now()
    };
    history[requestId] = result;
    writeTeamEmailHistory(history);
    return result;
  } finally {
    lock.releaseLock();
  }
}
