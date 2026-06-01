import { AUTH_TOKEN_SECRET_KEY, GOOGLE_CLIENT_ID_PROPERTY_KEY } from './constants';

declare const global: any;

const TOKEN_TTL = 5 * 60 * 1000; // 5 minutes
const GOOGLE_TOKENINFO_URL = 'https://oauth2.googleapis.com/tokeninfo?id_token=';

function ensureSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty(AUTH_TOKEN_SECRET_KEY);
  if (!secret) {
    secret = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid()));
    props.setProperty(AUTH_TOKEN_SECRET_KEY, secret);
  }
  return secret;
}

function sign(payload: string) {
  const secret = ensureSecret();
  const signature = Utilities.computeHmacSha256Signature(payload, secret);
  return Utilities.base64EncodeWebSafe(signature);
}

function encodePayload(data: Record<string, unknown>) {
  return Utilities.base64EncodeWebSafe(JSON.stringify(data));
}

function decodeToken(token: string) {
  try {
    const parts = token.split('.');
    if (parts.length !== 2) return null;
    const [payloadB64, sigB64] = parts;
    const expected = sign(payloadB64);
    if (sigB64 !== expected) return null;
    const payloadJson = Utilities.newBlob(Utilities.base64DecodeWebSafe(payloadB64)).getDataAsString();
    return JSON.parse(payloadJson);
  } catch (_) {
    return null;
  }
}

export function issueAuthToken() {
  const email = String(Session.getActiveUser?.().getEmail?.() || '').trim().toLowerCase();
  if (!email) throw new Error('Google sign-in required.');
  const now = Date.now();
  const payload = {
    email,
    iat: now,
    exp: now + TOKEN_TTL
  };
  const encoded = encodePayload(payload);
  const signature = sign(encoded);
  return `${encoded}.${signature}`;
}

export function getGoogleClientId() {
  return String(PropertiesService.getScriptProperties().getProperty(GOOGLE_CLIENT_ID_PROPERTY_KEY) || '').trim();
}

function decodeJwtPayload(token: string) {
  try {
    const parts = String(token || '').split('.');
    if (parts.length < 2) return null;
    const payloadJson = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString();
    return JSON.parse(payloadJson);
  } catch (_) {
    return null;
  }
}

function verifyGoogleIdToken(token: string) {
  const raw = String(token || '').trim();
  if (!raw) return null;

  const clientId = getGoogleClientId();
  if (!clientId) throw new Error('Google sign-in is not configured. Add GOOGLE_CLIENT_ID to Script Properties.');

  const cache = CacheService.getScriptCache();
  const payload = decodeJwtPayload(raw);
  const cacheKey = payload?.sub ? `google-id-token:${payload.sub}:${payload.exp || ''}` : '';
  if (cacheKey) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (_) { /* ignore */ }
    }
  }

  const response = UrlFetchApp.fetch(`${GOOGLE_TOKENINFO_URL}${encodeURIComponent(raw)}`, {
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) return null;

  const data = JSON.parse(response.getContentText() || '{}') as {
    aud?: string;
    email?: string;
    email_verified?: string | boolean;
    exp?: string;
    sub?: string;
  };

  if (String(data.aud || '').trim() !== clientId) return null;
  const email = String(data.email || '').trim().toLowerCase();
  if (!email) return null;

  const emailVerified = typeof data.email_verified === 'boolean'
    ? data.email_verified
    : String(data.email_verified || '').toLowerCase() === 'true';
  if (!emailVerified) return null;

  const verified = {
    email,
    sub: String(data.sub || '').trim(),
    expires: Number(data.exp || 0) * 1000
  };

  if (cacheKey) {
    const nowSec = Math.floor(Date.now() / 1000);
    const expSec = Number(data.exp || 0);
    const ttl = Math.max(1, Math.min(300, expSec - nowSec));
    cache.put(cacheKey, JSON.stringify(verified), ttl);
  }

  return verified;
}

export function verifyAuthToken(token?: string) {
  if (!token) return null;
  const payload = decodeToken(token);
  if (!payload || typeof payload.exp !== 'number' || typeof payload.email !== 'string') return null;
  if (Date.now() > payload.exp) return null;
  return {
    email: payload.email,
    issued: payload.iat,
    expires: payload.exp
  };
}

export function requestTokenEmail() {
  const raw = String(global.__REQUEST_AUTH_TOKEN__ || '').trim();
  if (!raw) return '';
  const data = verifyAuthToken(raw);
  if (data?.email) return data.email;
  const google = verifyGoogleIdToken(raw);
  return google?.email || '';
}
