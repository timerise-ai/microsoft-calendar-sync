# Data model

Five records, one store interface. Everything else in the skill is functions over these.

| Record | Key | Lives | Holds |
|---|---|---|---|
| `AuthRequest` | random 43-char `state` | 10 minutes, deleted on first read | user id, PKCE verifier |
| `CalendarConnection` | `userId` | until disconnect | account id and e-mail, status, write permission, mailbox zone, encrypted tokens |
| `LinkedCalendar` | `calendarRowId(userId, externalId)` | while Outlook lists the calendar | the two toggles, subscription pointer, sync state, lease |
| `WatchSubscription` | random UUID (carried in `clientState`) | until deleted or replaced | Graph's subscription id, expiry |
| `BusyBlock` | `busyBlockId(calendarId, eventId)` | until the event changes or leaves the window | `from`, `to` (UTC), all-day flag |

Booking events are not stored: the private marker on the event is the link, and Graph finds the event
by it. See [booking-events.md](booking-events.md).

## Rename table

Decide the host's words before copying, and rename everywhere at once.

| Skill | Source app called it | Typical host words |
|---|---|---|
| `userId` (the host) | user | staff, provider, employee, member, resource owner |
| `LinkedCalendar` | user calendar | connected calendar |
| `syncBusy` | `syncExceptions` | "check for conflicts", "block my availability" |
| `syncBookings` | `syncBookings` | "add bookings to calendar" |
| `BusyBlock` | exception (`slotType: 'EXCEPTION'`) | busy time, blocked time, conflict |
| `WatchSubscription` | calendar subscription | webhook subscription, push subscription |
| `BookingForCalendar` | booking | appointment, reservation, session, shift |

Keep Graph's own terms (`externalId`, `graphId`, `canEdit`, `showAs`, `clientState`) as they are;
they are the API's vocabulary, not the domain's.

**Running next to `google-calendar-sync`?** Both skills use the same record names on purpose. Keep them
in separate folders (`google-calendar/`, `microsoft-calendar/`) and separate collections or tables,
which the stores here already use (`microsoftCalendar…`, `microsoft_calendar_…`). The availability
engine reads both sets of busy blocks by user id and merges them.

## Lifecycle

```
             connect             toggle syncBusy on                       toggle off / calendar gone / disconnect
 (nothing) ──────────▶ rows ─────────────────────────▶ subscription + blocks ─────────────────────────────────────▶ (nothing)
                        │  invalid_grant / interaction_required      │  cron every 12 h: PATCH renew or replace, resync
                        ▼                                            ▼
                 reauth_required ── reconnect same account (tid:oid) ──▶ connected, settings kept, subscriptions renewed
                                 └─ reconnect other account ──────────▶ old rows torn down, new rows
```

## Types

```ts
// microsoft-calendar/types.ts
import type { GraphApi } from './graphApi';

/** One Microsoft account per user. `reauth_required` is set when the token endpoint refuses the refresh token. */
export type ConnectionStatus = 'connected' | 'reauth_required';

export interface CalendarConnection {
  userId: string;
  provider: 'microsoft';
  /** `tid:oid` from the id token: the directory and the user in it. Identifies the account on reconnect. */
  accountId: string;
  /** `preferred_username`, for display only: it can be renamed and is not unique across directories. */
  accountEmail: string;
  status: ConnectionStatus;
  /** False when only `Calendars.Read` was granted: busy sync works, booking sync does not. */
  canWrite: boolean;
  /** IANA zone of the mailbox. Anchors all-day events. Null without `MailboxSettings.Read`. */
  mailboxTimeZone: string | null;
  /** `StoredTokens` as JSON, AES-256-GCM encrypted. Never stored in clear text. */
  encryptedTokens: string;
  connectedAt: Date;
  updatedAt: Date;
  lastError: string | null;
}

export interface StoredTokens {
  accessToken: string;
  refreshToken: string;
  /** Epoch milliseconds. */
  expiresAt: number;
  scope: string;
}

/** A calendar of the connected account, as this app knows it. One row per (user, Outlook calendar). */
export interface LinkedCalendar {
  /** `calendarRowId(userId, externalId)`, deterministic so concurrent listings cannot create duplicates. */
  id: string;
  userId: string;
  /** Graph's calendar id, a long base64 string. */
  externalId: string;
  title: string;
  isDefault: boolean;
  isWritable: boolean;
  /** Inbound: this calendar's events block the user's availability. */
  syncBusy: boolean;
  /** Outbound: bookings hosted by this user are written to this calendar. */
  syncBookings: boolean;
  /** Our `WatchSubscription.id`, not Graph's. */
  subscriptionId: string | null;
  lastSyncedAt: Date | null;
  lastSyncError: string | null;
  syncLeaseUntil: Date | null;
  resyncRequested: boolean;
  /** Owned by the store: incremented on every write, used for compare-and-swap. */
  version: number;
}

export interface CalendarSettingsPatch {
  title?: string;
  syncBusy?: boolean;
  syncBookings?: boolean;
}

/** A Graph change-notification subscription on one calendar's events. */
export interface WatchSubscription {
  /** Our id, a UUID. It is what `clientState` carries, signed. */
  id: string;
  /** Graph's subscription id. Every notification must name it. */
  graphId: string;
  userId: string;
  calendarId: string;
  expiresAt: Date;
  createdAt: Date;
}

/** A busy interval imported from Outlook. Feeds the availability engine as host busy time. */
export interface BusyBlock {
  /** `busyBlockId(calendarId, eventId)`, deterministic so concurrent syncs converge instead of duplicating. */
  id: string;
  userId: string;
  calendarId: string;
  eventId: string;
  from: Date;
  to: Date;
  allDay: boolean;
}

/** Server-side state of an OAuth round trip, consumed exactly once. */
export interface AuthRequest {
  state: string;
  userId: string;
  codeVerifier: string;
  expiresAt: Date;
}

/** What the outbound side needs to know about a booking. The host maps its own booking onto this. */
export interface BookingForCalendar {
  bookingId: string;
  start: Date;
  end: Date;
  /** IANA zone the event is displayed in, usually the business's zone. */
  timeZone: string;
  title: string;
  guestName?: string;
  guestEmail?: string;
  /** Primary host first. Only the primary host's calendars receive the event. */
  hostIds: string[];
  hostEmails: Record<string, string | undefined>;
  links?: { label: string; url: string }[];
  detailsUrl?: string;
  notes?: string;
  wantsTeamsMeeting?: boolean;
}

export interface BookingSource {
  /** Confirmed bookings starting at or after `from` whose primary host is `userId`. */
  listUpcomingForPrimaryHost(userId: string, from: Date): Promise<BookingForCalendar[]>;
}

export interface Logger {
  info(kind: string, data?: Record<string, unknown>): void;
  warn(kind: string, data?: Record<string, unknown>): void;
  error(kind: string, data?: Record<string, unknown>): void;
}

export interface CalendarStore {
  saveAuthRequest(request: AuthRequest): Promise<void>;
  /** Read and delete in one atomic step. Expiry is checked by the caller. */
  consumeAuthRequest(state: string): Promise<AuthRequest | null>;
  deleteExpiredAuthRequests(before: Date): Promise<number>;

  getConnection(userId: string): Promise<CalendarConnection | null>;
  saveConnection(connection: CalendarConnection): Promise<void>;
  patchConnection(
    userId: string,
    patch: Partial<
      Pick<
        CalendarConnection,
        'status' | 'encryptedTokens' | 'updatedAt' | 'lastError' | 'canWrite' | 'mailboxTimeZone'
      >
    >
  ): Promise<void>;
  deleteConnection(userId: string): Promise<void>;

  listCalendars(userId: string): Promise<LinkedCalendar[]>;
  getCalendar(calendarId: string): Promise<LinkedCalendar | null>;
  /**
   * Compare-and-swap. `fn` receives the current row (or null) and returns the row to write, or null
   * to leave it unchanged. Retried on a concurrent write, so `fn` must be pure. Errors thrown by `fn`
   * propagate. The store sets `version`.
   */
  mutateCalendar(
    calendarId: string,
    fn: (current: LinkedCalendar | null) => LinkedCalendar | null
  ): Promise<{ before: LinkedCalendar | null; after: LinkedCalendar | null }>;
  deleteCalendar(calendarId: string): Promise<void>;
  /** Calendars with `syncBusy`, ordered by id, starting after `cursor`. */
  listBusySyncCalendars(cursor: string | null, limit: number): Promise<LinkedCalendar[]>;

  saveSubscription(subscription: WatchSubscription): Promise<void>;
  getSubscription(subscriptionId: string): Promise<WatchSubscription | null>;
  deleteSubscription(subscriptionId: string): Promise<void>;
  listSubscriptionsExpiringBefore(before: Date, limit: number): Promise<WatchSubscription[]>;

  listBusyBlocks(calendarId: string): Promise<BusyBlock[]>;
  applyBusyDiff(calendarId: string, upserts: BusyBlock[], deleteIds: string[]): Promise<void>;
  deleteBusyBlocks(calendarId: string): Promise<void>;
}

/** Who Outlook e-mails when a booking's event is written. Graph has no "send no invitations" switch. */
export type InviteAttendees = 'all' | 'hosts' | 'none';

export interface CalendarSyncConfig {
  clientId: string;
  clientSecret: string;
  /** `organizations` (work and school accounts), `common` (personal accounts too), or one tenant's id. */
  tenant: string;
  /** Exactly as registered on the app registration. */
  redirectUri: string;
  /** Public HTTPS URL of the notifications route. */
  notificationUrl: string;
  /** Also send lifecycle notifications (`subscriptionRemoved`, `missed`…) to `notificationUrl`. */
  lifecycleNotifications: boolean;
  /** HMAC key for `clientState`. Fails closed when empty. */
  clientStateSecret: string;
  /** 32 bytes as 64 hex chars or 44 base64 chars. */
  encryptionKey: string;
  /** GUID of the property set holding the booking marker. Generate one; keep it forever. */
  markerPropertySetId: string;
  /** All-day events are anchored here when neither the mailbox nor `hostTimeZone` gives a zone. */
  defaultTimeZone: string;
  /** How far ahead busy time is imported. Must cover the booking horizon. */
  horizonDays: number;
  /** Requested subscription lifetime. Graph's maximum for Outlook events is 10,080 minutes. */
  subscriptionLifetimeMs: number;
  /** Subscriptions expiring sooner than this are renewed by the cron. */
  renewBeforeMs: number;
  /** A crashed sync releases its lease after this long. */
  leaseMs: number;
  inviteAttendees: InviteAttendees;
}

export const DEFAULT_SYNC_SETTINGS = {
  tenant: 'organizations',
  lifecycleNotifications: true,
  horizonDays: 180,
  subscriptionLifetimeMs: 6 * 24 * 60 * 60 * 1000,
  renewBeforeMs: 48 * 60 * 60 * 1000,
  leaseMs: 2 * 60 * 1000,
  inviteAttendees: 'all',
} as const satisfies Partial<CalendarSyncConfig>;

export interface CalendarSyncDeps {
  store: CalendarStore;
  graph: GraphApi;
  config: CalendarSyncConfig;
  bookings: BookingSource;
  log: Logger;
  now: () => Date;
  /** Called with the Teams join URL Graph generated for a booking's event. */
  onConferenceLink?: (bookingId: string, url: string) => Promise<void>;
  /** Event title and body. Defaults to `defaultDescribeBooking`; supply one for i18n. */
  describeBooking?: (booking: BookingForCalendar) => { summary: string; description: string };
  /** Zone for all-day events when the mailbox zone is unknown, e.g. the host's business zone. */
  hostTimeZone?: (userId: string) => Promise<string | null>;
}

export type CalendarErrorCode =
  | 'not-connected'
  | 'reauth-required'
  | 'calendar-not-found'
  | 'calendar-not-writable'
  | 'invalid-or-expired-state'
  | 'no-refresh-token'
  | 'no-account-id'
  | 'missing-calendar-scope';

export class CalendarSyncError extends Error {
  readonly code: CalendarErrorCode;

  constructor(code: CalendarErrorCode, message: string = code) {
    super(message);
    this.name = 'CalendarSyncError';
    this.code = code;
  }
}
```

## Why the store looks like this

- **`mutateCalendar` is the only way to change a calendar row.** Settings toggles, the sync lease, the
  subscription pointer and error state all go through compare-and-swap on `version`. A toggle racing
  a sync, or two syncs racing each other, cannot overwrite one another's fields. `fn` may run several
  times, so it must be pure: no I/O inside, no counters outside that are not reset at its top (see the
  `acquired` flag in [busy-sync.md](busy-sync.md)).
- **`after` is the row as stored after the call.** When `fn` returns null nothing is written and
  `after` equals `before`. Callers rely on this to read the current state without a second query.
- **The sync lease lives on the calendar row**, not in a separate lock table. A crashed holder's lease
  simply expires after `leaseMs`.
- **Subscriptions are their own records keyed by our UUID**, because a notification arrives with only
  the clientState and Graph's subscription id, and because a subscription must outlive its calendar
  row until it has been deleted at Graph.
- **The account is `tid:oid`, not the e-mail.** A user can be renamed, and the same address can exist
  in two directories. The object id inside its tenant is what identifies the person on reconnect.
- **Busy blocks are per user, not per project.** A person is busy across every business they work for.
- **Tokens are one encrypted string.** The store never sees them in clear text, so a database export,
  backup or search-index sync cannot leak them.

## Checklist

- [ ] Rename decided and applied to types, store names, routes and UI copy
- [ ] `version` owned by the store; `fn` passed to `mutateCalendar` is pure
- [ ] No field of `StoredTokens` persisted outside `encryptedTokens`
- [ ] Busy blocks queried by user id from the availability engine
