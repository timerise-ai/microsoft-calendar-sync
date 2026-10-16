# OAuth and tokens

Connecting is a redirect round trip that ends with a refresh token. Four things decide whether it is
safe and whether it keeps working: who the callback attributes the tokens to, which account it was,
what was granted, and storing the *new* refresh token every time Microsoft hands one out.

## Entra ID app registration

| Step | Where | Value |
|---|---|---|
| Register | Entra admin center → App registrations → New registration | *Accounts in any organizational directory* for `tenant: 'organizations'`; add *and personal Microsoft accounts* for `common` |
| Redirect URI | Authentication → *Web* platform | `https://api.example.com/calendar/microsoft/callback`, exactly as `MICROSOFT_CALENDAR_REDIRECT_URI` |
| Client secret | Certificates & secrets | Expires (24 months at most): when it does, every connection fails. Calendar the rotation |
| Permissions | API permissions → Microsoft Graph → *Delegated* | `Calendars.ReadWrite`, `MailboxSettings.Read`, `offline_access`, `openid`, `profile` |
| Publisher | Branding & properties | Verify the publisher: many tenants block user consent to unverified multi-tenant apps |
| Marker GUID | your secret manager | `node -e "console.log(crypto.randomUUID())"` once, as `MICROSOFT_CALENDAR_MARKER_GUID`; never change it |

Register a separate app per environment. Graph allows one subscription per resource and app, so a
sandbox and a production deployment sharing one registration fight over each calendar's subscription.

**Tenants that disable user consent** show the user Microsoft's *Need admin approval* page. Send their
administrator the admin-consent link once:
`https://login.microsoftonline.com/{their-tenant}/adminconsent?client_id={clientId}`. After that, users
connect normally. Do not add `prompt=consent` to the sign-in URL; it forces the user-consent screen and
brings the approval page back even after an administrator consented.

## Scopes

| Scope | Gives | Needed for |
|---|---|---|
| `Calendars.ReadWrite` | list calendars, read events, subscribe, write events | everything; with only `Calendars.Read` booking sync is off |
| `MailboxSettings.Read` | the mailbox time zone | anchoring all-day events ([time-zones.md](time-zones.md)) |
| `offline_access` | a refresh token | anything after the first hour |
| `openid profile` | an id token with `tid`, `oid`, `preferred_username` | identifying the account on reconnect |

Read the token response's `scope` instead of assuming: an administrator can grant less than was
asked for. Graph returns some scopes as full URIs (`https://graph.microsoft.com/Calendars.ReadWrite`),
which `grantedScopes` normalises. Without a calendar scope the connection is refused
(`missing-calendar-scope`); with `Calendars.Read` only it is kept with `canWrite: false`; without
`MailboxSettings.Read` all-day events fall back to the host's zone.

## The round trip

```
GET /connect (signed in) ─▶ AuthRequest{state, userId, verifier, +10 min} ─▶ { url } with state + S256 challenge
Microsoft sign-in + consent ─▶ GET /callback?code&state ─▶ consume(state) atomically ─▶ exchange(code, verifier)
   ─▶ check scopes ─▶ id token tid:oid ─▶ mailbox zone ─▶ same account? keep : tear down ─▶ save
```

- **The state is a key, not a payload.** Anything put into it can be edited by whoever holds the URL;
  the callback must find the user server-side. Consuming it in a transaction makes it single use.
- **PKCE** binds the code to the server that started the flow.
- **`prompt=select_account`** lets someone signed in to several Microsoft accounts pick the right one.
- **The id token is read without checking its signature.** It came straight from the token endpoint
  over TLS, which OpenID Connect allows in place of signature validation. Never accept an id token
  from the browser this way.
- **Old subscriptions can only be deleted with the old account's tokens**, so tear-down runs before
  the new tokens replace the old.

## Token lifetime

| Event | What happens | Handled by |
|---|---|---|
| Access token near expiry (60–90 min) | refreshed 5 min early, saved encrypted | `withAccessToken` |
| A 401 on a call | one forced refresh and one retry | `withAccessToken` |
| Every refresh | **a new refresh token**, stored; the old one is not reused | `toStoredTokens` |
| `invalid_grant` on refresh | connection → `reauth_required` | `refreshTokens` |
| `interaction_required` on refresh | the same: MFA or Conditional Access now needs the user | `refreshTokens` |
| Reconnect, same `tid:oid` | `connected` again, settings kept, subscriptions renewed, resync | `completeConnect` |
| Disconnect | subscriptions deleted, blocks and rows deleted, tokens deleted | `disconnect` |

`invalid_grant` covers a refresh token unused for 90 days (AADSTS700082), revoked or invalidated by a
password reset (AADSTS50173), and consent withdrawn (AADSTS65001). A busy calendar is refreshed by every
cron run, so the 90-day expiry only catches hosts with booking sync alone and no bookings.

**Why not MSAL.** `@azure/msal-node` keeps refresh tokens in its own in-memory cache and does not
return them on its result. A server that stores MSAL's `AuthenticationResult` stores no refresh token;
when the first access token expires, every call fails until the user reconnects. Using MSAL properly
means persisting its whole token cache per user through a cache plugin. Two `fetch` calls to the token
endpoint are simpler and make `invalid_grant` visible.

Concurrent refreshes are harmless: a rotated-out refresh token stays valid until its own expiry, and
the last write wins.

**There is no revocation endpoint** for one app's refresh token. Disconnect forgets the tokens and
deletes the subscriptions; the grant stays listed for the user until they remove the app under
*My Apps* (work accounts) or *account.live.com/consent* (personal accounts).

## Encryption key

`CALENDAR_TOKEN_KEY` is 32 random bytes, as 64 hex or 44 base64 characters: `openssl rand -hex 32`.
Payloads are prefixed `v1:` so a second key can be introduced later.

## Code

```ts
// microsoft-calendar/crypto.ts
import {
  createCipheriv,
  createDecipheriv,
  createHash,
  createHmac,
  randomBytes,
  timingSafeEqual,
} from 'node:crypto';

const ALGORITHM = 'aes-256-gcm';
const IV_LENGTH = 12;
const TAG_LENGTH = 16;
const FORMAT_VERSION = 'v1';

/** Graph rejects a subscription whose `clientState` is longer than this. */
export const CLIENT_STATE_MAX_LENGTH = 128;

/** 32 bytes, given as 64 hex characters or 44 base64 characters. */
function parseKey(key: string): Buffer {
  if (key.length === 64 && /^[0-9a-fA-F]+$/.test(key)) return Buffer.from(key, 'hex');
  const buffer = Buffer.from(key, 'base64');
  if (buffer.length === 32) return buffer;
  throw new Error('Encryption key must be 32 bytes (64 hex or 44 base64 characters)');
}

/** `v1:base64(iv):base64(ciphertext):base64(tag)`. The version prefix leaves room for key rotation. */
export function encrypt(plaintext: string, key: string): string {
  const iv = randomBytes(IV_LENGTH);
  const cipher = createCipheriv(ALGORITHM, parseKey(key), iv, { authTagLength: TAG_LENGTH });
  const ciphertext = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  return [FORMAT_VERSION, iv, ciphertext, cipher.getAuthTag()]
    .map(part => (typeof part === 'string' ? part : part.toString('base64')))
    .join(':');
}

export function decrypt(payload: string, key: string): string {
  const [version, iv, ciphertext, tag] = payload.split(':');
  if (version !== FORMAT_VERSION || !iv || !ciphertext || !tag) {
    throw new Error('Unrecognised encrypted payload');
  }
  const decipher = createDecipheriv(ALGORITHM, parseKey(key), Buffer.from(iv, 'base64'), {
    authTagLength: TAG_LENGTH,
  });
  decipher.setAuthTag(Buffer.from(tag, 'base64'));
  return Buffer.concat([decipher.update(Buffer.from(ciphertext, 'base64')), decipher.final()]).toString(
    'utf8'
  );
}

export function randomToken(bytes = 32): string {
  return randomBytes(bytes).toString('base64url');
}

/** PKCE S256 challenge for a verifier. */
export function pkceChallenge(verifier: string): string {
  return createHash('sha256').update(verifier).digest('base64url');
}

export function sha256Hex(input: string): string {
  return createHash('sha256').update(input).digest('hex');
}

/** Deterministic row id: listing the same calendar twice, concurrently, writes one row. */
export function calendarRowId(userId: string, externalId: string): string {
  return sha256Hex(`calendar\n${userId}\n${externalId}`).slice(0, 40);
}

/** Deterministic busy-block id. Graph event ids contain `/`, `+` and `=`; hashing keeps ids path-safe. */
export function busyBlockId(calendarId: string, eventId: string): string {
  return sha256Hex(`busy\n${calendarId}\n${eventId}`).slice(0, 40);
}

function hmac(data: string, secret: string): string {
  return createHmac('sha256', secret).update(data).digest('base64url');
}

/**
 * The `clientState` Graph echoes in every notification for a subscription. It carries only our
 * subscription id (a UUID): 36 + 1 + 43 = 80 characters, well inside Graph's 128. Anything else,
 * such as a calendar id (150+ characters), would make Graph refuse the subscription.
 */
export function signClientState(subscriptionId: string, secret: string): string {
  if (!secret) throw new Error('clientState secret is not configured');
  const value = `${subscriptionId}.${hmac(`subscription:${subscriptionId}`, secret)}`;
  if (value.length > CLIENT_STATE_MAX_LENGTH) throw new Error('clientState exceeds 128 characters');
  return value;
}

/** The subscription id the clientState was issued for, or null. */
export function verifyClientState(clientState: string, secret: string): string | null {
  if (!secret) return null;
  const dot = clientState.lastIndexOf('.');
  if (dot <= 0) return null;
  const subscriptionId = clientState.slice(0, dot);
  const received = Buffer.from(clientState.slice(dot + 1));
  const expected = Buffer.from(hmac(`subscription:${subscriptionId}`, secret));
  // Both are base64url SHA-256 digests (43 chars); a length mismatch is already a forgery.
  if (received.length !== expected.length || !timingSafeEqual(received, expected)) return null;
  return subscriptionId;
}

/** Constant-time comparison for bearer secrets such as the cron secret. */
export function safeEqual(a: string, b: string): boolean {
  const left = createHash('sha256').update(a).digest();
  const right = createHash('sha256').update(b).digest();
  return timingSafeEqual(left, right) && a.length === b.length;
}
```

```ts
// microsoft-calendar/tokens.ts
import { decrypt, encrypt } from './crypto';
import { GraphApiError, type TokenResponse } from './graphApi';
import { CalendarSyncError, type CalendarSyncDeps, type StoredTokens } from './types';

/** Refresh this long before expiry, so a token never expires between check and use. */
const EXPIRY_SKEW_MS = 5 * 60_000;

/**
 * Token-endpoint errors that only a new sign-in fixes. `invalid_grant`: revoked, expired after
 * 90 days unused, password reset, consent withdrawn. `interaction_required`: MFA or a Conditional
 * Access policy now demands the user.
 */
const REAUTH_CODES = new Set(['invalid_grant', 'interaction_required']);

export function sealTokens(deps: CalendarSyncDeps, tokens: StoredTokens): string {
  return encrypt(JSON.stringify(tokens), deps.config.encryptionKey);
}

export function openTokens(deps: CalendarSyncDeps, encryptedTokens: string): StoredTokens {
  return JSON.parse(decrypt(encryptedTokens, deps.config.encryptionKey)) as StoredTokens;
}

export function toStoredTokens(
  response: TokenResponse,
  now: Date,
  previous?: StoredTokens
): StoredTokens | null {
  // Microsoft rotates the refresh token on every refresh: keep the new one. Fall back to the old
  // one only if a response ever omits it.
  const refreshToken = response.refresh_token ?? previous?.refreshToken;
  if (!refreshToken) return null;
  return {
    accessToken: response.access_token,
    refreshToken,
    expiresAt: now.getTime() + response.expires_in * 1000,
    scope: response.scope ?? previous?.scope ?? '',
  };
}

/** Scope names as granted, lower-cased and without the `https://graph.microsoft.com/` prefix. */
export function grantedScopes(scope: string): Set<string> {
  return new Set(
    scope
      .split(' ')
      .filter(Boolean)
      .map(s => s.replace(/^https:\/\/graph\.microsoft\.com\//i, '').toLowerCase())
  );
}

async function refreshTokens(
  deps: CalendarSyncDeps,
  userId: string,
  current: StoredTokens
): Promise<StoredTokens> {
  let response: TokenResponse;
  try {
    response = await deps.graph.refresh(current.refreshToken);
  } catch (error) {
    if (error instanceof GraphApiError && REAUTH_CODES.has(error.code)) {
      await deps.store.patchConnection(userId, {
        status: 'reauth_required',
        lastError: error.message,
        updatedAt: deps.now(),
      });
      deps.log.warn('microsoft-calendar-reauth-required', { userId, code: error.code });
      throw new CalendarSyncError('reauth-required');
    }
    throw error;
  }
  const next = toStoredTokens(response, deps.now(), current);
  if (!next) throw new CalendarSyncError('no-refresh-token');
  await deps.store.patchConnection(userId, {
    encryptedTokens: sealTokens(deps, next),
    updatedAt: deps.now(),
  });
  return next;
}

/**
 * Run `fn` with a valid access token for the user's Microsoft account. Refreshes ahead of expiry and
 * once more on a 401 (a token revoked early, or a clock skew), then gives up.
 */
export async function withAccessToken<T>(
  deps: CalendarSyncDeps,
  userId: string,
  fn: (accessToken: string) => Promise<T>
): Promise<T> {
  const connection = await deps.store.getConnection(userId);
  if (!connection) throw new CalendarSyncError('not-connected');
  if (connection.status === 'reauth_required') throw new CalendarSyncError('reauth-required');

  let tokens = openTokens(deps, connection.encryptedTokens);
  if (tokens.expiresAt - EXPIRY_SKEW_MS <= deps.now().getTime()) {
    tokens = await refreshTokens(deps, userId, tokens);
  }
  try {
    return await fn(tokens.accessToken);
  } catch (error) {
    if (!(error instanceof GraphApiError) || error.status !== 401) throw error;
    tokens = await refreshTokens(deps, userId, tokens);
    return fn(tokens.accessToken);
  }
}
```

```ts
// microsoft-calendar/connect.ts
import { syncBusy } from './busySync';
import { refreshCalendarList } from './calendars';
import { pkceChallenge, randomToken } from './crypto';
import { ensureSubscription, teardownCalendar } from './subscriptions';
import { toIanaTimeZone } from './timeZones';
import { grantedScopes, sealTokens, toStoredTokens } from './tokens';
import { CalendarSyncError, type CalendarConnection, type CalendarSyncDeps } from './types';

const AUTH_REQUEST_TTL_MS = 10 * 60 * 1000;
// randomToken(32) is 43 base64url characters. Checked before the lookup: the state arrives in the
// query string and becomes a document id, which must not contain '/'.
const STATE_PATTERN = /^[A-Za-z0-9_-]{43}$/;

/**
 * The account behind a token response, from its id token. The token came straight from the token
 * endpoint over TLS, which OpenID Connect accepts in place of checking the signature.
 */
export function accountFromIdToken(idToken: string | undefined): { accountId: string; accountEmail: string } | null {
  const payload = idToken?.split('.')[1];
  if (!payload) return null;
  try {
    const claims = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8')) as Record<string, unknown>;
    const { tid, oid } = claims;
    if (typeof tid !== 'string' || typeof oid !== 'string') return null;
    const email = claims.preferred_username ?? claims.email;
    return { accountId: `${tid}:${oid}`, accountEmail: typeof email === 'string' ? email : '' };
  } catch {
    return null;
  }
}

/**
 * Begin a connection for the signed-in user. The state is an unguessable key of a server-side
 * record holding the user id and the PKCE verifier; the callback never reads a user id from the URL.
 */
export async function startConnect(deps: CalendarSyncDeps, userId: string): Promise<string> {
  const state = randomToken(32);
  const codeVerifier = randomToken(48);
  await deps.store.saveAuthRequest({
    state,
    userId,
    codeVerifier,
    expiresAt: new Date(deps.now().getTime() + AUTH_REQUEST_TTL_MS),
  });
  return deps.graph.authUrl({ state, codeChallenge: pkceChallenge(codeVerifier) });
}

/**
 * Finish the OAuth round trip. Reconnecting the same Microsoft account keeps every calendar's
 * settings and renews its subscriptions (the usual case: fixing `reauth_required`). Connecting a
 * different account tears the old one down first, because its calendar ids mean nothing to the new tokens.
 */
export async function completeConnect(
  deps: CalendarSyncDeps,
  params: { state: string; code: string }
): Promise<{ userId: string; sameAccount: boolean }> {
  const request = STATE_PATTERN.test(params.state)
    ? await deps.store.consumeAuthRequest(params.state)
    : null;
  if (!request || request.expiresAt <= deps.now()) throw new CalendarSyncError('invalid-or-expired-state');
  const { userId } = request;

  const response = await deps.graph.exchangeCode(params.code, request.codeVerifier);
  const tokens = toStoredTokens(response, deps.now());
  if (!tokens) throw new CalendarSyncError('no-refresh-token');

  // An administrator may have granted less than was asked for; the response says what was granted.
  const granted = grantedScopes(tokens.scope);
  const canWrite = granted.has('calendars.readwrite');
  if (!canWrite && !granted.has('calendars.read')) throw new CalendarSyncError('missing-calendar-scope');

  const account = accountFromIdToken(response.id_token);
  if (!account) throw new CalendarSyncError('no-account-id');

  let mailboxTimeZone: string | null = null;
  if (granted.has('mailboxsettings.read') || granted.has('mailboxsettings.readwrite')) {
    try {
      mailboxTimeZone = toIanaTimeZone(await deps.graph.mailboxTimeZone(tokens.accessToken));
    } catch (error) {
      deps.log.info('microsoft-calendar-mailbox-zone-unavailable', {
        userId,
        message: error instanceof Error ? error.message : String(error),
      });
    }
  }

  const previous = await deps.store.getConnection(userId);
  const sameAccount = previous?.accountId === account.accountId;

  // Old subscriptions can only be deleted with the old account's tokens, so this runs before they are replaced.
  if (previous && !sameAccount) {
    for (const calendar of await deps.store.listCalendars(userId)) {
      await teardownCalendar(deps, calendar);
    }
  }

  const now = deps.now();
  const connection: CalendarConnection = {
    userId,
    provider: 'microsoft',
    accountId: account.accountId,
    accountEmail: account.accountEmail,
    status: 'connected',
    canWrite,
    mailboxTimeZone: mailboxTimeZone ?? (sameAccount ? (previous?.mailboxTimeZone ?? null) : null),
    encryptedTokens: sealTokens(deps, tokens),
    connectedAt: sameAccount && previous ? previous.connectedAt : now,
    updatedAt: now,
    lastError: null,
  };
  await deps.store.saveConnection(connection);

  if (sameAccount) {
    // Rows pick up the new write permission; busy calendars get their subscription renewed (or
    // replaced, if Graph dropped it while the tokens were dead) and a resync.
    const calendars = await refreshCalendarList(deps, userId);
    for (const calendar of calendars.filter(c => c.syncBusy)) {
      try {
        await ensureSubscription(deps, calendar, { force: true });
        await syncBusy(deps, calendar.id);
      } catch (error) {
        deps.log.warn('microsoft-calendar-resume-failed', {
          calendarId: calendar.id,
          message: error instanceof Error ? error.message : String(error),
        });
      }
    }
  }
  deps.log.info('microsoft-calendar-connected', { userId, sameAccount, canWrite, mailboxTimeZone });
  return { userId, sameAccount };
}

/**
 * Delete every subscription, busy block and row, and forget the tokens. Microsoft has no endpoint
 * that revokes one app's refresh token: the grant stays listed for the user until they remove the
 * app under My Apps (work accounts) or account.live.com/consent (personal accounts). The refresh
 * token itself expires after 90 days unused.
 */
export async function disconnect(deps: CalendarSyncDeps, userId: string): Promise<void> {
  const connection = await deps.store.getConnection(userId);
  if (!connection) return;

  for (const calendar of await deps.store.listCalendars(userId)) {
    await teardownCalendar(deps, calendar);
  }
  await deps.store.deleteConnection(userId);
  deps.log.info('microsoft-calendar-disconnected', { userId });
}
```

## Checklist

- [ ] One app registration per environment; redirect URI identical in Entra and config
- [ ] Publisher verified; admin-consent link ready for tenants that need it
- [ ] User id read from the consumed `AuthRequest`, never from the callback URL
- [ ] Every token response's `refresh_token` stored, encrypted
- [ ] `reauth_required` shown in the UI with a reconnect button
- [ ] Client secret expiry on someone's calendar
