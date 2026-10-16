# Firestore backend

`CalendarStore` on `firebase-admin`. Five top-level collections, all server-only.

| Collection | Doc id | Fields |
|---|---|---|
| `microsoftCalendarAuthRequests` | `state` | `userId`, `codeVerifier`, `expiresAt` |
| `microsoftCalendarConnections` | `userId` | `accountId`, `accountEmail`, `status`, `canWrite`, `mailboxTimeZone`, `encryptedTokens`, `connectedAt`, `updatedAt`, `lastError` |
| `microsoftCalendarLinks` | `calendarRowId(userId, externalId)` | the `LinkedCalendar` fields |
| `microsoftCalendarSubscriptions` | subscription UUID | `graphId`, `userId`, `calendarId`, `expiresAt`, `createdAt` |
| `microsoftCalendarBusyBlocks` | `busyBlockId(calendarId, eventId)` | `userId`, `calendarId`, `eventId`, `from`, `to`, `allDay` |

Top-level rather than `users/{uid}/...` subcollections: the cron pages across all users' calendars, the
webhook resolves a subscription without knowing its user, and the availability query reads blocks by
user id across calendars. All three are one query on a top-level collection. The `microsoftCalendar`
prefix keeps them apart from `google-calendar-sync`'s collections in the same project.

## Indexes

| Collection | Fields | Used by |
|---|---|---|
| `microsoftCalendarBusyBlocks` | `userId` ASC, `to` ASC | `listHostBusy` (availability) |

Every other query uses one equality filter, or one range filter ordered by the same field, which
Firestore's automatic single-field indexes serve. The `firestore.indexes.json` entry:

```json
{
  "indexes": [
    {
      "collectionGroup": "microsoftCalendarBusyBlocks",
      "queryScope": "COLLECTION",
      "fields": [
        { "fieldPath": "userId", "order": "ASCENDING" },
        { "fieldPath": "to", "order": "ASCENDING" }
      ]
    }
  ]
}
```

## Security rules

Nothing here is for clients. The Admin SDK bypasses rules; deny everything else, inside the existing
`match /databases/{database}/documents` block:

```
match /microsoftCalendarAuthRequests/{id} { allow read, write: if false; }
match /microsoftCalendarConnections/{id} { allow read, write: if false; }
match /microsoftCalendarLinks/{id} { allow read, write: if false; }
match /microsoftCalendarSubscriptions/{id} { allow read, write: if false; }
match /microsoftCalendarBusyBlocks/{id} { allow read, write: if false; }
```

## Notes

- **`mutateCalendar` is a transaction.** Firestore retries it on contention, which is why the function
  passed in must be pure.
- **`consumeAuthRequest` deletes inside the transaction whether or not the request expired**, so an
  expired state cannot be retried until it happens to be valid.
- **Batches cap at 500 writes**; `applyBusyDiff` and the deletes chunk accordingly. That is safe because
  the diff is idempotent: an interrupted run is completed by the next one.
- **Dates are `Timestamp` in Firestore and `Date` everywhere else**; `toDoc`/`fromDoc` convert at the
  boundary so no `Timestamp` leaks into the domain functions.
- **A `/` in a document id throws.** Subscription ids come from verified clientStates, but
  `getSubscription` still refuses one rather than turning a bad request into a 500. Outlook ids, which
  do contain `/`, are never document ids: rows and blocks use hashes.
- **TTL policy (optional):** a Firestore TTL policy on `microsoftCalendarAuthRequests.expiresAt` removes
  abandoned requests without the cron's sweep.

## Code

```ts
// microsoft-calendar/store/firestore.ts
import {
  FieldPath,
  Timestamp,
  type DocumentData,
  type DocumentReference,
  type Firestore,
  type WriteBatch,
} from 'firebase-admin/firestore';

import type {
  AuthRequest,
  BusyBlock,
  CalendarConnection,
  CalendarStore,
  LinkedCalendar,
  WatchSubscription,
} from '../types';

const C = {
  authRequests: 'microsoftCalendarAuthRequests',
  connections: 'microsoftCalendarConnections',
  calendars: 'microsoftCalendarLinks',
  subscriptions: 'microsoftCalendarSubscriptions',
  busy: 'microsoftCalendarBusyBlocks',
} as const;

const BATCH_LIMIT = 500;

/** Timestamp to Date for the listed fields; everything else passes through. */
function fromDoc<T>(data: DocumentData, dateFields: readonly string[]): T {
  const out: Record<string, unknown> = { ...data };
  for (const field of dateFields) {
    const value = data[field];
    out[field] = value instanceof Timestamp ? value.toDate() : (value ?? null);
  }
  return out as T;
}

function toDoc(data: object): DocumentData {
  const out: DocumentData = {};
  for (const [key, value] of Object.entries(data)) {
    if (value === undefined) continue;
    out[key] = value instanceof Date ? Timestamp.fromDate(value) : value;
  }
  return out;
}

const CALENDAR_DATES = ['lastSyncedAt', 'syncLeaseUntil'] as const;
const SUBSCRIPTION_DATES = ['expiresAt', 'createdAt'] as const;
const BUSY_DATES = ['from', 'to'] as const;

async function deleteInBatches(db: Firestore, refs: DocumentReference[]): Promise<void> {
  for (let i = 0; i < refs.length; i += BATCH_LIMIT) {
    const batch = db.batch();
    for (const ref of refs.slice(i, i + BATCH_LIMIT)) batch.delete(ref);
    await batch.commit();
  }
}

export function createFirestoreCalendarStore(db: Firestore): CalendarStore {
  const calendars = db.collection(C.calendars);
  const subscriptions = db.collection(C.subscriptions);
  const busy = db.collection(C.busy);

  return {
    async saveAuthRequest(request) {
      await db.collection(C.authRequests).doc(request.state).set(toDoc(request));
    },

    consumeAuthRequest(state) {
      const ref = db.collection(C.authRequests).doc(state);
      return db.runTransaction(async tx => {
        const snapshot = await tx.get(ref);
        if (!snapshot.exists) return null;
        // Deleted whether or not it is still valid, so a replay finds nothing.
        tx.delete(ref);
        return fromDoc<AuthRequest>(snapshot.data() ?? {}, ['expiresAt']);
      });
    },

    async deleteExpiredAuthRequests(before) {
      const snapshot = await db
        .collection(C.authRequests)
        .where('expiresAt', '<', Timestamp.fromDate(before))
        .limit(BATCH_LIMIT)
        .get();
      await deleteInBatches(db, snapshot.docs.map(doc => doc.ref));
      return snapshot.size;
    },

    async getConnection(userId) {
      const snapshot = await db.collection(C.connections).doc(userId).get();
      return snapshot.exists
        ? fromDoc<CalendarConnection>(snapshot.data() ?? {}, ['connectedAt', 'updatedAt'])
        : null;
    },

    async saveConnection(connection) {
      await db.collection(C.connections).doc(connection.userId).set(toDoc(connection));
    },

    async patchConnection(userId, patch) {
      await db.collection(C.connections).doc(userId).update(toDoc(patch));
    },

    async deleteConnection(userId) {
      await db.collection(C.connections).doc(userId).delete();
    },

    async listCalendars(userId) {
      const snapshot = await calendars.where('userId', '==', userId).get();
      return snapshot.docs.map(doc => fromDoc<LinkedCalendar>(doc.data(), CALENDAR_DATES));
    },

    async getCalendar(calendarId) {
      const snapshot = await calendars.doc(calendarId).get();
      return snapshot.exists ? fromDoc<LinkedCalendar>(snapshot.data() ?? {}, CALENDAR_DATES) : null;
    },

    mutateCalendar(calendarId, fn) {
      const ref = calendars.doc(calendarId);
      // Firestore retries the transaction on contention, which is why `fn` must be pure.
      return db.runTransaction(async tx => {
        const snapshot = await tx.get(ref);
        const before = snapshot.exists ? fromDoc<LinkedCalendar>(snapshot.data() ?? {}, CALENDAR_DATES) : null;
        const next = fn(before);
        if (!next) return { before, after: before };
        const after: LinkedCalendar = { ...next, id: calendarId, version: (before?.version ?? 0) + 1 };
        tx.set(ref, toDoc(after));
        return { before, after };
      });
    },

    async deleteCalendar(calendarId) {
      await calendars.doc(calendarId).delete();
    },

    async listBusySyncCalendars(cursor, limit) {
      let query = calendars.where('syncBusy', '==', true).orderBy(FieldPath.documentId()).limit(limit);
      if (cursor) query = query.startAfter(cursor);
      const snapshot = await query.get();
      return snapshot.docs.map(doc => fromDoc<LinkedCalendar>(doc.data(), CALENDAR_DATES));
    },

    async saveSubscription(subscription) {
      await subscriptions.doc(subscription.id).set(toDoc(subscription));
    },

    async getSubscription(subscriptionId) {
      // Ids come from verified clientStates, but a '/' would still make doc() throw.
      if (!subscriptionId || subscriptionId.includes('/')) return null;
      const snapshot = await subscriptions.doc(subscriptionId).get();
      return snapshot.exists ? fromDoc<WatchSubscription>(snapshot.data() ?? {}, SUBSCRIPTION_DATES) : null;
    },

    async deleteSubscription(subscriptionId) {
      await subscriptions.doc(subscriptionId).delete();
    },

    async listSubscriptionsExpiringBefore(before, limit) {
      const snapshot = await subscriptions
        .where('expiresAt', '<', Timestamp.fromDate(before))
        .orderBy('expiresAt')
        .limit(limit)
        .get();
      return snapshot.docs.map(doc => fromDoc<WatchSubscription>(doc.data(), SUBSCRIPTION_DATES));
    },

    async listBusyBlocks(calendarId) {
      const snapshot = await busy.where('calendarId', '==', calendarId).get();
      return snapshot.docs.map(doc => fromDoc<BusyBlock>(doc.data(), BUSY_DATES));
    },

    async applyBusyDiff(_calendarId, upserts, deleteIds) {
      const writes: ((batch: WriteBatch) => void)[] = [
        ...upserts.map(block => (batch: WriteBatch) => batch.set(busy.doc(block.id), toDoc(block))),
        ...deleteIds.map(id => (batch: WriteBatch) => batch.delete(busy.doc(id))),
      ];
      for (let i = 0; i < writes.length; i += BATCH_LIMIT) {
        const batch = db.batch();
        for (const write of writes.slice(i, i + BATCH_LIMIT)) write(batch);
        await batch.commit();
      }
    },

    async deleteBusyBlocks(calendarId) {
      const snapshot = await busy.where('calendarId', '==', calendarId).select().get();
      await deleteInBatches(db, snapshot.docs.map(doc => doc.ref));
    },
  };
}

/**
 * Host busy time for the availability engine: every block of the user overlapping [from, to).
 * Needs the composite index `microsoftCalendarBusyBlocks (userId ASC, to ASC)`.
 */
export async function listHostBusy(
  db: Firestore,
  userId: string,
  range: { from: Date; to: Date }
): Promise<{ from: Date; to: Date }[]> {
  const snapshot = await db
    .collection(C.busy)
    .where('userId', '==', userId)
    .where('to', '>', Timestamp.fromDate(range.from))
    .get();
  return snapshot.docs
    .map(doc => fromDoc<BusyBlock>(doc.data(), BUSY_DATES))
    .filter(block => block.from < range.to)
    .map(block => ({ from: block.from, to: block.to }));
}
```

## Checklist

- [ ] Composite index `microsoftCalendarBusyBlocks (userId, to)` deployed
- [ ] Rules deny client access to all five collections
- [ ] The availability engine reads host busy time through `listHostBusy` or an equivalent query
