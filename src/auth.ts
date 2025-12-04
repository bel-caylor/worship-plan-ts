import { AUTH_TOKEN_SECRET_KEY } from './constants';

declare const global: any;

const TOKEN_TTL = 5 * 60 * 1000; // 5 minutes

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
  return data ? data.email : '';
}
