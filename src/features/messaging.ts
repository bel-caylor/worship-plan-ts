import { getViewerProfile } from './roles';

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
