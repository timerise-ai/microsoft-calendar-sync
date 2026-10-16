# Subscriptions

A Graph subscription is Microsoft's promise to POST to your URL when a calendar's events change. It
expires, Graph can remove it, and nothing renews it for you. This file covers opening, renewing,
replacing and deleting subscriptions, and the cron that keeps them, and the busy blocks, current.

## Lifecycle

| Moment | Action | Failure handling |
|---|---|---|
| Busy sync switched on | `openSubscription`, then first `syncBusy` | The sync runs anyway; `lastSyncError` recorded; the cron opens it |
| Every cron run, no record | open one | Error recorded, resync still runs; retried next run |
| Every cron run, expiring within `renewBeforeMs` | PATCH a new expiry | Old one kept; retried next run |
| PATCH answers 404 | Graph dropped it: delete the record, open a new one | Same |
| POST answers 409 | Graph still holds one for this calendar: delete it, POST again | Same |
| Every cron run | `syncBusy` on every busy calendar | Recorded on the row |
| Lifecycle `subscriptionRemoved` / `reauthorizationRequired` | renew or replace, then resync | Logged; the cron retries |
| Lifecycle `missed` | resync | Logged |
| Busy sync switched off | row unlinked, subscription deleted, blocks deleted | Delete is best effort; the record always goes |
| Calendar gone, disconnect, other account | `teardownCalendar` | Same |
| Subscription no row points at | deleted by the cron | Orphans from races or crashes do not live forever |

- **A subscription failure never stops the sync.** Opening or renewing runs first, so a change made
  during the first sync still arrives, but its error is held until the sync has run. The other order
  means a calendar Graph will not subscribe (a bad notification URL, a clientState it refuses) never
  imports any busy time, while its toggle says "on".
- **`clientState` carries our subscription id only, signed.** Graph caps it at 128 characters and
  refuses the subscription above that. Our UUID plus an HMAC is 80; a calendar id alone is around 150.
  Everything else is looked up from the stored record.
- **Renew in place, never new-before-old.** Graph rejects a second subscription with the same resource
  and change types (409), so the Google pattern of opening a new channel before stopping the old one
  does not apply. A PATCH keeps the subscription id and the clientState; there is no gap to bridge.
- **404 on renewal is a fact, not an error to retry.** The subscription expired (the cron did not run
  for days) or Graph removed it (access revoked, mailbox moved). Retrying the PATCH forever leaves the
  calendar without notifications while every check of the local record says it has one.
- **409 on create is an orphan.** A subscription Graph holds but we have no record of, left by a crash
  between the POST and saving the record, or by a record deleted by hand. Its clientState is
  unknown, so it cannot be adopted; delete it and create ours.
- **Read the expiry Graph returns.** Outlook event subscriptions may last up to 10,080 minutes (seven
  days). The default asks for six; with a 48-hour renewal window and a 12-hourly cron, each subscription
  gets four chances to renew.
- **Lifecycle notifications** go to the same URL (`lifecycleNotificationUrl = notificationUrl`). Graph
  sends `subscriptionRemoved` when it drops a subscription, `reauthorizationRequired` when its access
  needs refreshing (renewing does that), and `missed` when it could not deliver. Each one triggers the
  repair it names, so a dropped subscription is replaced within minutes instead of at the next cron.

## Code

```ts
// microsoft-calendar/subscriptions.ts
import { randomUUID } from 'node:crypto';

import { syncBusy } from './busySync';
import { signClientState } from './crypto';
import { GraphApiError, type GraphApi, type GraphSubscriptionResource } from './graphApi';
import { withAccessToken } from './tokens';
import type { CalendarSyncDeps, LinkedCalendar, WatchSubscription } from './types';

/** The resource a calendar's subscription watches. Graph allows one subscription per resource and change type. */
export function subscriptionResource(externalId: string): string {
  return `/me/calendars/${encodeURIComponent(externalId)}/events`;
}

const sameResource = (a: string, b: string) =>
  decodeURIComponent(a).replace(/^\//, '').toLowerCase() === decodeURIComponent(b).replace(/^\//, '').toLowerCase();

async function createAtGraph(
  graph: GraphApi,
  accessToken: string,
  body: Parameters<GraphApi['createSubscription']>[1]
): Promise<GraphSubscriptionResource> {
  try {
    return await graph.createSubscription(accessToken, body);
  } catch (error) {
    // 409: Graph already holds a subscription on this resource for this app, left behind by a crash
    // or a lost record. It cannot be adopted (its clientState is unknown), so delete it and retry once.
    if (!(error instanceof GraphApiError) || error.status !== 409) throw error;
    for (const existing of await graph.listSubscriptions(accessToken)) {
      if (sameResource(existing.resource, body.resource)) await graph.deleteSubscription(accessToken, existing.id);
    }
    return graph.createSubscription(accessToken, body);
  }
}

/**
 * Subscribe to the calendar's events and point the row at the subscription. Returns null when busy
 * sync was switched off while the subscription was being created (it is deleted again rather than
 * left delivering for a calendar nobody syncs).
 */
export async function openSubscription(
  deps: CalendarSyncDeps,
  calendar: LinkedCalendar
): Promise<WatchSubscription | null> {
  const id = randomUUID();
  const { notificationUrl, lifecycleNotifications, subscriptionLifetimeMs, clientStateSecret } = deps.config;
  const created = await withAccessToken(deps, calendar.userId, accessToken =>
    createAtGraph(deps.graph, accessToken, {
      changeType: 'created,updated,deleted',
      resource: subscriptionResource(calendar.externalId),
      notificationUrl,
      ...(lifecycleNotifications ? { lifecycleNotificationUrl: notificationUrl } : {}),
      expirationDateTime: new Date(deps.now().getTime() + subscriptionLifetimeMs).toISOString(),
      clientState: signClientState(id, clientStateSecret),
    })
  );
  const subscription: WatchSubscription = {
    id,
    graphId: created.id,
    userId: calendar.userId,
    calendarId: calendar.id,
    // Read the expiry Graph grants; it may shorten what was asked for.
    expiresAt: new Date(created.expirationDateTime),
    createdAt: deps.now(),
  };
  await deps.store.saveSubscription(subscription);

  const { after } = await deps.store.mutateCalendar(calendar.id, current =>
    current?.syncBusy ? { ...current, subscriptionId: id } : null
  );
  if (after?.subscriptionId !== id) {
    await closeSubscription(deps, subscription);
    return null;
  }
  return subscription;
}

/** Best effort at Graph, always locally. A subscription Graph forgot about costs nothing. */
export async function closeSubscription(deps: CalendarSyncDeps, subscription: WatchSubscription): Promise<void> {
  try {
    await withAccessToken(deps, subscription.userId, accessToken =>
      deps.graph.deleteSubscription(accessToken, subscription.graphId)
    );
  } catch (error) {
    if (!(error instanceof GraphApiError) || error.status !== 404) {
      deps.log.warn('microsoft-calendar-subscription-delete-failed', {
        subscriptionId: subscription.id,
        message: error instanceof Error ? error.message : String(error),
      });
    }
  }
  await deps.store.deleteSubscription(subscription.id);
}

/**
 * Make sure the calendar has a live subscription. Renews in place (PATCH) when it expires within
 * `renewBeforeMs`, or always with `force`. A subscription Graph no longer has (404: it expired,
 * or Graph removed it) is replaced by a new one instead of being retried forever.
 */
export async function ensureSubscription(
  deps: CalendarSyncDeps,
  calendar: LinkedCalendar,
  options: { force?: boolean } = {}
): Promise<{ subscription: WatchSubscription | null; action: 'kept' | 'renewed' | 'opened' }> {
  const current = calendar.subscriptionId ? await deps.store.getSubscription(calendar.subscriptionId) : null;
  if (!current) return { subscription: await openSubscription(deps, calendar), action: 'opened' };

  const renewBy = deps.now().getTime() + deps.config.renewBeforeMs;
  if (!options.force && current.expiresAt.getTime() > renewBy) return { subscription: current, action: 'kept' };

  try {
    const renewed = await withAccessToken(deps, calendar.userId, accessToken =>
      deps.graph.renewSubscription(
        accessToken,
        current.graphId,
        new Date(deps.now().getTime() + deps.config.subscriptionLifetimeMs)
      )
    );
    const subscription = { ...current, expiresAt: new Date(renewed.expirationDateTime) };
    await deps.store.saveSubscription(subscription);
    return { subscription, action: 'renewed' };
  } catch (error) {
    if (!(error instanceof GraphApiError) || error.status !== 404) throw error;
    await deps.store.deleteSubscription(current.id);
    return { subscription: await openSubscription(deps, calendar), action: 'opened' };
  }
}

export async function enableBusySync(deps: CalendarSyncDeps, calendar: LinkedCalendar): Promise<void> {
  // Subscription first: a change made while the first sync runs then still arrives as a notification.
  // A subscription that cannot be opened must not stop the sync: busy time is imported now, the
  // error is recorded, and the cron keeps trying to subscribe.
  let failure: unknown = null;
  try {
    await openSubscription(deps, calendar);
  } catch (error) {
    failure = error;
  }
  await syncBusy(deps, calendar.id);
  if (failure) throw failure;
}

export async function disableBusySync(deps: CalendarSyncDeps, calendar: LinkedCalendar): Promise<void> {
  const { before } = await deps.store.mutateCalendar(calendar.id, current =>
    current ? { ...current, subscriptionId: null, syncLeaseUntil: null, resyncRequested: false } : null
  );
  const subscription = before?.subscriptionId ? await deps.store.getSubscription(before.subscriptionId) : null;
  if (subscription) await closeSubscription(deps, subscription);
  await deps.store.deleteBusyBlocks(calendar.id);
}

/** Delete the subscription, the busy blocks and the row. Used on disconnect and when Outlook drops a calendar. */
export async function teardownCalendar(deps: CalendarSyncDeps, calendar: LinkedCalendar): Promise<void> {
  const subscription = calendar.subscriptionId ? await deps.store.getSubscription(calendar.subscriptionId) : null;
  if (subscription) await closeSubscription(deps, subscription);
  await deps.store.deleteBusyBlocks(calendar.id);
  await deps.store.deleteCalendar(calendar.id);
}

async function mapPool<T>(items: T[], size: number, fn: (item: T) => Promise<void>): Promise<void> {
  let next = 0;
  const workers = Array.from({ length: Math.min(size, items.length) }, async () => {
    while (next < items.length) {
      const item = items[next++] as T;
      await fn(item);
    }
  });
  await Promise.all(workers);
}

export interface RenewReport {
  calendars: number;
  renewed: number;
  opened: number;
  synced: number;
  failed: number;
  orphansDeleted: number;
}

/**
 * The cron. For every calendar with busy sync: open a subscription if it has none, renew it in place
 * if it expires within `renewBeforeMs`, replace it if Graph lost it; then resync, which moves the
 * horizon forward and repairs any notification that was lost. Finally deletes subscriptions no
 * calendar points at. A failure is recorded on the row and retried on the next run.
 */
export async function renewSubscriptions(deps: CalendarSyncDeps, concurrency = 4): Promise<RenewReport> {
  const report: RenewReport = { calendars: 0, renewed: 0, opened: 0, synced: 0, failed: 0, orphansDeleted: 0 };

  let cursor: string | null = null;
  for (;;) {
    const page = await deps.store.listBusySyncCalendars(cursor, 100);
    if (page.length === 0) break;
    cursor = page[page.length - 1]?.id ?? null;
    report.calendars += page.length;

    await mapPool(page, concurrency, async calendar => {
      try {
        // As in enableBusySync: the resync runs even when the subscription could not be kept.
        let failure: unknown = null;
        try {
          const { action } = await ensureSubscription(deps, calendar);
          if (action === 'renewed') report.renewed++;
          if (action === 'opened') report.opened++;
        } catch (error) {
          failure = error;
        }
        const outcome = await syncBusy(deps, calendar.id);
        if (outcome.status === 'synced') report.synced++;
        if (failure) throw failure;
      } catch (error) {
        report.failed++;
        const message = error instanceof Error ? error.message : String(error);
        await deps.store.mutateCalendar(calendar.id, row => (row ? { ...row, lastSyncError: message } : null));
        deps.log.error('microsoft-calendar-renew-failed', { calendarId: calendar.id, message });
      }
    });
  }

  const renewBy = new Date(deps.now().getTime() + deps.config.renewBeforeMs);
  for (const subscription of await deps.store.listSubscriptionsExpiringBefore(renewBy, 500)) {
    const calendar = await deps.store.getCalendar(subscription.calendarId);
    if (calendar?.syncBusy && calendar.subscriptionId === subscription.id) continue;
    await closeSubscription(deps, subscription);
    report.orphansDeleted++;
  }

  await deps.store.deleteExpiredAuthRequests(deps.now());
  deps.log.info('microsoft-calendar-renewed', { ...report });
  return report;
}
```

## Checklist

- [ ] Cron scheduled every 12 h (or more often than `renewBeforeMs / 2`)
- [ ] One app registration per environment, so environments do not 409 each other
- [ ] Renewal is a PATCH; a 404 replaces, a 409 clears and retries
- [ ] Busy-sync toggle failures recorded, not reverted
