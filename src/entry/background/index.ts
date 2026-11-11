// MSRC Helper — background service worker.
//
// Three responsibilities:
//
// 1. Session seeding. The portal (msrc.microsoft.com) caches its MSAL tokens
//    in sessionStorage, which Chrome partitions per tab: a new tab has no
//    tokens and gets bounced to an interactive B2C login. Any logged-in
//    portal tab publishes its msal.* snapshot; this worker caches it in
//    chrome.storage.session (memory-only) and serves it to fresh tabs before
//    the SPA boots.
//
// 2. Awaiting computation. "Awaiting" = reports whose latest thread message
//    is the researcher's own — a dimension the portal UI does not offer. The
//    list-page chip and the popup dashboard both consume it. These calls are
//    proxied through a portal tab's content script (see src/entry/proxy)
//    because the APIs reject extension-origin requests.
//
// 3. Notification badge. The refresh token is persisted in
//    chrome.storage.local so it survives browser restarts (the same trick
//    mail-checker extensions use); an alarm mints fresh access tokens from
//    it every 15 minutes and polls the unread-notification count, shown on
//    the toolbar icon. Works with no portal tab open.

const SNAPSHOT_KEY = "msalSnapshot";
const PERSISTENT_SNAPSHOT_KEY = "msalSnapshotPersistent";
const MSRC_ORIGIN = "https://msrc.microsoft.com";
const MSRC_URLS = [MSRC_ORIGIN + "/*"];
const MESSAGE_API = "https://api.msrc.microsoft.com/message";
const NOTIFICATION_API = "https://api.msrc.microsoft.com/portal/v2.0/notification";
const B2C_TOKEN_ENDPOINT =
  "https://msrcweb.b2clogin.com/msrcweb.onmicrosoft.com/b2c_1a_multitenant_signupsignin/oauth2/v2.0/token";
const B2C_CLIENT_ID = "1d13e529-99b8-48fe-a77d-844853ae25bf";
const B2C_SCOPE = "https://msrcweb.onmicrosoft.com/msrc-api/user_impersonation openid profile";
const CHUNK = 40; // ids per `in (...)` query, keeps URLs short
const PAGE = 100; // the API rejects $top above 100
const POLL_ALARM = "msrc-notify-poll";

type Snapshot = Record<string, string>;

interface PortalMessage {
  reportNumber: string;
  senderName: string;
  createdOn: string;
  body?: string;
}

interface PortalReport {
  reportNumber: string;
  externalId?: string;
  title?: string;
  status?: string;
  ticketNumber?: string;
  securityImpact?: string;
  lastActivityOn?: string;
}

interface AwaitingCase {
  vuln: string;
  title: string;
  status: string;
  waitedSince: string; // ISO timestamp of the researcher's latest message
}

// Top-level code runs on every service-worker start. Content scripts are
// "untrusted contexts" and can only read storage.session after the access
// level is widened — do it before anything else.
chrome.storage.session.setAccessLevel({
  accessLevel: "TRUSTED_AND_UNTRUSTED_CONTEXTS",
});

/* ---------------- session seeding ---------------- */

async function readSnapshot(): Promise<Snapshot | undefined> {
  const stored = await chrome.storage.session.get(SNAPSHOT_KEY);
  return stored[SNAPSHOT_KEY] as Snapshot | undefined;
}

// Fallback when the cached snapshot is missing: ask live portal tabs directly.
async function liveTabSnapshot(): Promise<Snapshot | undefined> {
  const tabs = await chrome.tabs.query({ url: MSRC_URLS });
  for (const tab of tabs) {
    if (tab.id === undefined) continue;
    try {
      const snap = (await chrome.tabs.sendMessage(tab.id, {
        type: "querySnapshot",
      })) as Snapshot | undefined;
      if (snap && Object.keys(snap).length > 0) return snap;
    } catch {
      // tab not answering (content script not injected yet) — try the next
    }
  }
  return undefined;
}

// Defense in depth: snapshots may only come from our own content script
// running on the MSRC portal. Runtime messages are extension-internal
// anyway; this guards against future wiring mistakes.
function isMsrcSender(sender: chrome.runtime.MessageSender): boolean {
  try {
    return !!sender.url && new URL(sender.url).origin === MSRC_ORIGIN;
  } catch {
    return false;
  }
}

/* ---------------- portal API access (page-origin proxy) ---------------- */

// Every API call goes through a portal tab: its isolated-world script relays
// to the MAIN-world proxy, whose XHR carries the page origin. The requesting
// tab is recorded here before each compute job (single-flight, so plain
// module state is safe) and tried first — its message port is alive by
// definition, which keeps this working right after extension reloads, when
// other tabs may still host orphaned content scripts with dead ports.
let proxyTabId: number | undefined;

async function portalApiGet<T>(url: string, attempts = 3, body?: string): Promise<T> {
  const tabs = await chrome.tabs.query({ url: MSRC_URLS });
  const ordered = tabs
    .filter((t) => t.id !== undefined && t.id === proxyTabId)
    .concat(tabs.filter((t) => t.id !== proxyTabId));
  let lastError: unknown = null;
  for (const tab of ordered) {
    if (tab.id === undefined) continue;
    // The portal sheds requests intermittently; one transient XHR error
    // must not kill a whole multi-page compute — retry each tab a few
    // times before moving on to the next one. Callers doing best-effort
    // work (enrichment) pass attempts=1 and fail fast instead.
    for (let attempt = 0; attempt < attempts; attempt++) {
      if (attempt > 0) {
        const { promise: backoff, resolve: wake } = Promise.withResolvers<void>();
        setTimeout(wake, 400 * attempt);
        await backoff;
      }
      try {
        const resp = (await chrome.tabs.sendMessage(tab.id, {
          type: "apiGet",
          url,
          body,
        })) as {
          status?: number;
          body?: T;
          error?: string;
        };
        if (resp?.error) throw new Error(resp.error);
        if (!resp || resp.status !== 200) throw new Error(`api ${resp?.status}`);
        return resp.body as T;
      } catch (e) {
        lastError = e;
      }
    }
  }
  throw new Error(`no portal tab answered: ${String(lastError)}`);
}

async function pagedMessages(filter: string): Promise<PortalMessage[]> {
  const all: PortalMessage[] = [];
  for (let skip = 0; skip < 500; skip += PAGE) {
    const url =
      `${MESSAGE_API}?$orderBy=createdOn%20desc` +
      `&$filter=${encodeURIComponent(filter)}&$top=${PAGE}&$skip=${skip}`;
    const body = await portalApiGet<{ value?: PortalMessage[] }>(url);
    const value = body.value || [];
    all.push(...value);
    if (value.length < PAGE) break; // last page
  }
  return all;
}

// The REST /portal/v2.0/vulnerabilityReport endpoint 403s page-origin
// requests (its CORS story changes day to day), but the GraphQL endpoint
// serves the same data reliably. Filter shape and the or-chain batching
// (the `in` operator trips the server's field-cost limit; or-chains of eq
// at 20 ids do not) were verified against the live API.
const GRAPHQL_API = "https://api.msrc.microsoft.com/portalgraphql/v2.0";
const GQL_BATCH = 20;

async function fetchReports(ids: string[]): Promise<PortalReport[]> {
  const out: PortalReport[] = [];
  for (let i = 0; i < ids.length; i += GQL_BATCH) {
    const batch = ids.slice(i, i + GQL_BATCH);
    const alternatives = batch
      .map((id) => `{ reportNumber: { eq: "${id}" } }`)
      .join(",");
    const query =
      `query { vulnerabilityReports(where: { or: [${alternatives}] }, take: ${batch.length})` +
      ` { items { reportNumber title status externalId ticketNumber securityImpact lastActivityOn } } }`;
    const resp = await portalApiGet<{ data?: { vulnerabilityReports?: { items?: PortalReport[] } }; errors?: Array<{ message: string }> }>(
      GRAPHQL_API,
      1, // best effort — enrichment must not break the awaiting set
      JSON.stringify({ query })
    );
    if (resp?.errors?.length) throw new Error(resp.errors[0].message);
    out.push(...(resp?.data?.vulnerabilityReports?.items || []));
  }
  return out;
}

// Run `fn` over the id list in chunks so `in (...)` URLs stay short.
async function inChunks<T>(ids: string[], fn: (chunk: string[]) => Promise<T[]>): Promise<T[]> {
  const out: T[] = [];
  for (let i = 0; i < ids.length; i += CHUNK) {
    out.push(...(await fn(ids.slice(i, i + CHUNK))));
  }
  return out;
}

/* ---------------- awaiting computation ---------------- */

// Note: the API silently ignores `senderName ne 'MSRC'` filters, so "my
// messages" queries actually return everyone's. That is fine here — both
// steps below only need per-report message ordering, not filtering.
async function computeAwaiting(): Promise<{ cases: AwaitingCase[]; cves: Record<string, string> }> {
  // Step 1: newest messages first, across all reports that have any thread.
  // The same pages feed two consumers: per-report message ordering below,
  // and CVE extraction (MSRC delivers assigned CVE numbers inside ordinary
  // thread messages — there is no structured field for them).
  const mine = await pagedMessages("conversationId eq null");
  const lastMine = new Map<string, string>(); // reportNumber -> createdOn of its newest message
  for (const m of mine) {
    if (!lastMine.has(m.reportNumber)) lastMine.set(m.reportNumber, m.createdOn);
  }
  if (lastMine.size === 0) return { cases: [], cves: {} };

  // CVE extraction: collect per-report ids, then drop boilerplate. Message
  // templates cite the same CVE across many threads (e.g. a policy example);
  // an id that shows up on more than three different reports is noise, not
  // an assignment.
  const perReport = new Map<string, string[]>();
  const reportCount = new Map<string, number>();
  for (const m of mine) {
    const found = (m.body || "").match(/CVE-\d{4}-\d{4,7}/gi) || [];
    if (found.length === 0) continue;
    const seen = perReport.get(m.reportNumber) || [];
    for (const raw of found) {
      const cve = raw.toUpperCase();
      if (seen.includes(cve)) continue; // already counted for this report
      seen.push(cve);
      reportCount.set(cve, (reportCount.get(cve) || 0) + 1);
    }
    perReport.set(m.reportNumber, seen);
  }
  const cves: Record<string, string> = {};
  for (const [rn, list] of perReport) {
    const real = list.filter((cve) => (reportCount.get(cve) || 0) <= 3);
    if (real.length > 0) cves[rn] = real.join(", ");
  }

  // Step 2: the sender of each report's newest message decides direction —
  // anything other than "MSRC" means the researcher spoke last.
  const rns = Array.from(lastMine.keys());
  const msgs = await inChunks(rns, (chunk) =>
    pagedMessages(`conversationId eq null and reportNumber in ('${chunk.join("','")}')`)
  );
  const latest = new Map<string, PortalMessage>();
  for (const m of msgs) {
    if (!latest.has(m.reportNumber)) latest.set(m.reportNumber, m);
  }
  const waiting = Array.from(latest.entries())
    .filter(([, message]) => message.senderName !== "MSRC")
    .map(([rn]) => rn);
  if (waiting.length === 0) return { cases: [], cves };

  // Step 3: enrich with title/status for the popup dashboard and panel.
  // Best effort via GraphQL — titles are cosmetic: the waiting ids and wait
  // times from steps 1-2 are already complete.
  try {
    const reps = await fetchReports(waiting);
    const byRn = new Map(reps.map((r) => [r.reportNumber, r]));
    return {
      cases: waiting.map((rn) => ({
        vuln: rn,
        title: byRn.get(rn)?.title || "",
        status: byRn.get(rn)?.status || "",
        waitedSince: lastMine.get(rn) || "",
      })),
      cves,
    };
  } catch (e) {
    console.warn("MSRC helper: report enrichment unavailable, listing ids only", e);
    return {
      cases: waiting.map((rn) => ({
        vuln: rn,
        title: "",
        status: "",
        waitedSince: lastMine.get(rn) || "",
      })),
      cves,
    };
  }
}

/* ---------------- awaiting cache ---------------- */

// The service worker is killed after ~30s idle, and module state dies with
// it. The cache is therefore mirrored into chrome.storage.session (memory
// only, cleared when the browser exits), so a fresh worker — and the chip's
// first paint after a page refresh — still sees the last known count
// instead of a misleading zero.
const AWAITING_CACHE_KEY = "awaitingCache";

const awaiting = {
  cases: [] as AwaitingCase[],
  cves: {} as Record<string, string>, // reportNumber -> extracted CVE ids
  fetchedAt: 0,
  newestKey: "", // identity of the newest thread message at last compute
  inflight: null as Promise<AwaitingCase[]> | null,
  loaded: false, // has storage state been read into memory?
};

async function loadAwaitingFromStorage(): Promise<void> {
  if (awaiting.loaded) return;
  awaiting.loaded = true;
  try {
    const stored = (await chrome.storage.session.get(AWAITING_CACHE_KEY))[AWAITING_CACHE_KEY] as
      | { cases: AwaitingCase[]; cves?: Record<string, string>; fetchedAt: number; newestKey: string }
      | undefined;
    if (stored && stored.fetchedAt > 0) {
      awaiting.cases = stored.cases;
      awaiting.cves = stored.cves || {};
      awaiting.fetchedAt = stored.fetchedAt;
      awaiting.newestKey = stored.newestKey;
    }
  } catch {
    // storage unavailable — start cold
  }
}

async function saveAwaitingToStorage(): Promise<void> {
  try {
    await chrome.storage.session.set({
      [AWAITING_CACHE_KEY]: {
        cases: awaiting.cases,
        cves: awaiting.cves,
        fetchedAt: awaiting.fetchedAt,
        newestKey: awaiting.newestKey,
      },
    });
  } catch {
    // best effort — an unpersisted cache only costs one recompute
  }
}

async function cachedAwaiting() {
  await loadAwaitingFromStorage();
  return { cases: awaiting.cases, ageMs: Date.now() - awaiting.fetchedAt };
}

// The newest thread message of ANY sender fingerprints the whole computation:
// a researcher reply (starts a wait) or an MSRC reply (ends one) both move
// it. Verified against the live API: plain $top=1, no $select (this backend
// 400s on unverified query options).
async function newestMessageKey(): Promise<string> {
  const url =
    `${MESSAGE_API}?$orderBy=createdOn%20desc` +
    `&$filter=${encodeURIComponent("conversationId eq null")}&$top=1`;
  const body = await portalApiGet<{ value?: Array<{ id: string; createdOn: string }> }>(url);
  const first = (body.value || [])[0];
  return first ? `${first.id}|${first.createdOn}` : "";
}

function getAwaiting(requestingTabId?: number): Promise<AwaitingCase[]> {
  if (awaiting.inflight) return awaiting.inflight;
  const job = (async () => {
    await loadAwaitingFromStorage();
    proxyTabId = requestingTabId; // preferred proxy host for this job
    let key = "";
    try {
      key = await newestMessageKey();
    } catch {
      // fingerprint probe failed — fall through to a full recompute
    }
    if (awaiting.fetchedAt > 0 && key !== "" && key === awaiting.newestKey) {
      awaiting.fetchedAt = Date.now();
      return awaiting.cases; // nothing changed since the last compute
    }
    const result = await computeAwaiting();
    awaiting.cases = result.cases;
    awaiting.cves = result.cves;
    // An empty result is more likely a flaky fetch than a researcher with
    // zero open questions — never mark it fresh, so the next request
    // recomputes instead of serving it forever.
    const trustworthy = key !== "" && result.cases.length > 0;
    awaiting.fetchedAt = trustworthy ? Date.now() : 0;
    awaiting.newestKey = key;
    void saveAwaitingToStorage();
    return result.cases;
  })();
  awaiting.inflight = job;
  job.catch(() => undefined).finally(() => {
    awaiting.inflight = null;
  });
  return job;
}

/* ---------------- message router ---------------- */

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  const fromPortalPage = isMsrcSender(sender);
  const fromOwnExtension = sender.id === chrome.runtime.id;
  void (async () => {
    switch (msg?.type) {
      case "getSnapshot": {
        if (!fromPortalPage) {
          sendResponse(undefined);
          break;
        }
        let snap = await readSnapshot();
        if (!snap || Object.keys(snap).length === 0) {
          // session storage dies with the browser; the disk copy is what
          // keeps a new tab logged in after a full restart. The access
          // token inside may be expired — MSAL renews it from the RT.
          snap = await readPersistentSnapshot();
          if (snap && Object.keys(snap).length > 0) {
            await chrome.storage.session.set({ [SNAPSHOT_KEY]: snap });
          }
        }
        if (!snap || Object.keys(snap).length === 0) {
          snap = await liveTabSnapshot();
        }
        if (!snap || Object.keys(snap).length === 0) {
          // the source tab may still be re-injecting after an extension
          // reload — give it one short second chance
          const { promise: retryTick, resolve: tick } = Promise.withResolvers<void>();
          setTimeout(tick, 300);
          await retryTick;
          snap = await liveTabSnapshot();
        }
        sendResponse(snap && Object.keys(snap).length > 0 ? snap : undefined);
        break;
      }
      case "pushSnapshot": {
        if (!fromPortalPage) {
          sendResponse(false);
          break;
        }
        const entries = msg.entries as Snapshot;
        if (entries && Object.keys(entries).length > 0) {
          await chrome.storage.session.set({ [SNAPSHOT_KEY]: entries });
          await persistSnapshot(entries); // refresh token at rest for the badge
          void pollNotifications(); // fresh login — update the badge now
        } else {
          // every pushing tab is logged out — drop the cached sessions
          await chrome.storage.session.remove(SNAPSHOT_KEY);
          await chrome.storage.local.remove(PERSISTENT_SNAPSHOT_KEY);
          await chrome.action.setBadgeText({ text: "" });
        }
        sendResponse(true);
        break;
      }
      case "getAwaitingCached": {
        // instant cache-only read for first-paint rendering
        if (!fromOwnExtension) {
          sendResponse({ error: "forbidden" });
          break;
        }
        void cachedAwaiting().then(sendResponse);
        break;
      }
      case "getReportMeta": {
        // short MSRC case ids (externalId) for the list-page column
        if (!fromOwnExtension) {
          sendResponse({ error: "forbidden" });
          break;
        }
        try {
          const ids = (msg.ids as string[]) || [];
          const reps = await fetchReports(ids);
          const records: Record<string, { ext: string; title: string; status: string; ticket: string; impact: string }> = {};
          for (const r of reps) {
            records[r.reportNumber] = {
              ext: r.externalId || "",
              title: r.title || "",
              status: r.status || "",
              ticket: r.ticketNumber || "",
              impact: r.securityImpact || "",
            };
          }
          await loadAwaitingFromStorage(); // cves ride the awaiting cache
          sendResponse({ records, cves: awaiting.cves });
        } catch (e) {
          sendResponse({ error: String(e) });
        }
        break;
      }
      case "getAwaiting": {
        if (!fromOwnExtension) {
          sendResponse({ error: "forbidden" });
          break;
        }
        try {
          sendResponse({ cases: await getAwaiting(sender.tab?.id) });
        } catch (e) {
          sendResponse({ error: String(e) });
        }
        break;
      }
      case "getNewMessages": {
        if (!fromOwnExtension) {
          sendResponse({ error: "forbidden" });
          break;
        }
        try {
          sendResponse({ cases: await loadNewMessages() });
        } catch (e) {
          sendResponse({ error: String(e) });
        }
        break;
      }
      default:
        sendResponse(undefined);
    }
  })();
  return true; // keep the channel open for the async sendResponse
});

/* ---------------- notification badge ---------------- */

// The portal APIs reject chrome-extension:// origins server-side, but the
// host_permissions above let the worker reach them; a DNR rule rewrites the
// Origin header to the portal's so the server accepts the request. This is
// what lets the badge poll run with no portal tab open.
const DNR_RULE_IDS = [100, 101];

async function installOriginRules(): Promise<void> {
  const rules: chrome.declarativeNetRequest.Rule[] = [
    {
      id: DNR_RULE_IDS[0],
      priority: 1,
      action: {
        type: "modifyHeaders" as chrome.declarativeNetRequest.RuleActionType,
        requestHeaders: [
          {
            header: "Origin",
            operation: "set" as chrome.declarativeNetRequest.HeaderOperation,
            value: MSRC_ORIGIN,
          },
        ],
      },
      condition: {
        urlFilter: "||api.msrc.microsoft.com/",
        // no resourceTypes: a service-worker fetch is often "other", and an
        // xmlhttprequest-only rule would never rewrite Origin
        initiatorDomains: [chrome.runtime.id],
      },
    },
    {
      id: DNR_RULE_IDS[1],
      priority: 1,
      action: {
        type: "modifyHeaders" as chrome.declarativeNetRequest.RuleActionType,
        requestHeaders: [
          {
            header: "Origin",
            operation: "set" as chrome.declarativeNetRequest.HeaderOperation,
            value: MSRC_ORIGIN,
          },
        ],
      },
      condition: {
        urlFilter: "||msrcweb.b2clogin.com/",
        initiatorDomains: [chrome.runtime.id],
      },
    },
  ];
  try {
    await chrome.declarativeNetRequest.updateDynamicRules({
      removeRuleIds: DNR_RULE_IDS,
      addRules: rules,
    });
  } catch (e) {
    console.warn("MSRC helper: DNR rules unavailable", e);
  }
}

// The persistent snapshot (chrome.storage.local, survives restarts) holds the
// refresh token so the badge keeps working after a browser restart — the same
// approach mail-checker extensions use. Deliberate trade-off: the RT is at
// rest in the profile's extension storage; it never leaves the machine.
async function readPersistentSnapshot(): Promise<Snapshot | undefined> {
  const stored = await chrome.storage.local.get(PERSISTENT_SNAPSHOT_KEY);
  return stored[PERSISTENT_SNAPSHOT_KEY] as Snapshot | undefined;
}

async function persistSnapshot(entries: Snapshot): Promise<void> {
  try {
    await chrome.storage.local.set({ [PERSISTENT_SNAPSHOT_KEY]: entries });
  } catch {
    // best effort — the badge simply pauses until the next portal login
  }
}

function refreshTokenFrom(snapshot: Snapshot | undefined): string | null {
  if (!snapshot) return null;
  // Match the key, not the secret shape: B2C refresh tokens are often "eyJ"
  // too, so a "0." prefix heuristic returns null and the badge never polls.
  for (const [key, value] of Object.entries(snapshot)) {
    if (!key.includes("|refreshtoken|")) continue;
    try {
      const parsed: unknown = JSON.parse(value);
      if (parsed && typeof parsed === "object" && "secret" in parsed) {
        const secret = (parsed as { secret: unknown }).secret;
        if (typeof secret === "string" && secret.length > 20) return secret;
      }
    } catch {
      // not a JSON value — skip
    }
  }
  return null;
}

let mintedToken: { token: string; expiresAt: number } | null = null;

// B2C rotates the refresh token on use. The response's new RT must replace
// the one at rest, or the next poll (and every poll after a restart) fails.
async function persistRotatedRefreshToken(rotated: string): Promise<void> {
  const snapshot = await readPersistentSnapshot();
  if (!snapshot) return;
  const next: Snapshot = { ...snapshot };
  let changed = false;
  for (const [key, value] of Object.entries(snapshot)) {
    if (!key.includes("|refreshtoken|")) continue;
    try {
      const parsed: unknown = JSON.parse(value);
      if (!parsed || typeof parsed !== "object" || !("secret" in parsed)) continue;
      const updated = { ...parsed, secret: rotated };
      next[key] = JSON.stringify(updated);
      changed = true;
    } catch {
      // leave the entry untouched
    }
  }
  if (!changed) return;
  await persistSnapshot(next);
  await chrome.storage.session.set({ [SNAPSHOT_KEY]: next });
}

async function mintAccessToken(): Promise<string | null> {
  if (mintedToken && mintedToken.expiresAt > Date.now() + 60_000) {
    return mintedToken.token;
  }
  const rt = refreshTokenFrom(await readPersistentSnapshot());
  if (!rt) return null;
  const body =
    `client_id=${B2C_CLIENT_ID}&grant_type=refresh_token` +
    `&refresh_token=${encodeURIComponent(rt)}&scope=${encodeURIComponent(B2C_SCOPE)}`;
  const resp = await fetch(B2C_TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  if (!resp.ok) {
    mintedToken = null; // RT expired — badge pauses until the next portal login
    return null;
  }
  const raw: unknown = await resp.json();
  if (!raw || typeof raw !== "object") return null;
  const access = "access_token" in raw ? raw.access_token : undefined;
  if (typeof access !== "string" || !access) return null;
  const expiresIn = "expires_in" in raw && typeof raw.expires_in === "number" ? raw.expires_in : 3600;
  const rotated = "refresh_token" in raw && typeof raw.refresh_token === "string" ? raw.refresh_token : null;
  if (rotated && rotated !== rt) await persistRotatedRefreshToken(rotated);
  mintedToken = { token: access, expiresAt: Date.now() + expiresIn * 1000 };
  return mintedToken.token;
}

interface NewMessageCase {
  vuln: string;
  extId: string;
  title: string;
  status: string;
  createdOn: string;
}

// Service-worker fetch against the message API. The DNR rules rewrite Origin
// so this works without a portal tab, like the badge poll.
async function swPagedMessages(filter: string, maxPages = 20): Promise<PortalMessage[]> {
  const token = await mintAccessToken();
  if (!token) throw new Error("no session — open the portal once");
  const all: PortalMessage[] = [];
  for (let skip = 0; skip < maxPages * PAGE; skip += PAGE) {
    const url =
      `${MESSAGE_API}?$orderBy=createdOn%20desc` +
      `&$filter=${encodeURIComponent(filter)}&$top=${PAGE}&$skip=${skip}`;
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
    });
    if (!resp.ok) throw new Error(`api ${resp.status}`);
    const raw: unknown = await resp.json();
    if (!raw || typeof raw !== "object" || !("value" in raw) || !Array.isArray(raw.value)) break;
    const page = raw.value.filter(
      (m): m is PortalMessage =>
        !!m && typeof m === "object" && "reportNumber" in m && typeof m.reportNumber === "string"
    );
    all.push(...page);
    if (page.length < PAGE) break;
  }
  return all;
}

// Popup list: a case has a new message when MSRC sent its latest thread
// message — the exact opposite of the Waiting set, from the message API
// itself. The notification feed's read flag proved unreliable (it reported
// zero while threads had fresh MSRC replies), so it is not used here. The
// toolbar badge still counts unread notifications.
async function loadNewMessages(): Promise<NewMessageCase[]> {
  const token = await mintAccessToken();
  if (!token) throw new Error("no session — open the portal once");
  const messages = await swPagedMessages("conversationId eq null");
  const latest = new Map<string, PortalMessage>();
  for (const m of messages) {
    if (!latest.has(m.reportNumber)) latest.set(m.reportNumber, m);
  }
  const fresh = Array.from(latest.entries()).filter(
    ([, message]) => message.senderName === "MSRC"
  );
  fresh.sort((a, b) => Date.parse(b[1].createdOn) - Date.parse(a[1].createdOn));
  const cases: NewMessageCase[] = fresh.map(([rn, message]) => ({
    vuln: rn,
    extId: "",
    title: "",
    status: "",
    createdOn: message.createdOn,
  }));
  if (cases.length === 0) return cases;
  try {
    // same or-eq GraphQL batch as the grid columns, but over a direct
    // service-worker fetch (no portal tab needed)
    for (let i = 0; i < cases.length; i += GQL_BATCH) {
      const batch = cases.slice(i, i + GQL_BATCH);
      const alternatives = batch
        .map((c) => `{ reportNumber: { eq: "${c.vuln}" } }`)
        .join(",");
      const query =
        `query { vulnerabilityReports(where: { or: [${alternatives}] }, take: ${batch.length})` +
        ` { items { reportNumber title status externalId } } }`;
      const resp = await fetch(GRAPHQL_API, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
          Accept: "application/json",
        },
        body: JSON.stringify({ query }),
      });
      if (!resp.ok) continue;
      const raw: unknown = await resp.json();
      if (!raw || typeof raw !== "object" || !("data" in raw)) continue;
      const data = raw.data;
      if (!data || typeof data !== "object" || !("vulnerabilityReports" in data)) continue;
      const reports = data.vulnerabilityReports;
      if (!reports || typeof reports !== "object" || !("items" in reports)) continue;
      const items = reports.items;
      if (!Array.isArray(items)) continue;
      const byRn = new Map<string, PortalReport>();
      for (const item of items) {
        if (!item || typeof item !== "object") continue;
        if (!("reportNumber" in item) || typeof item.reportNumber !== "string") continue;
        byRn.set(item.reportNumber, item as PortalReport);
      }
      for (const c of batch) {
        const report = byRn.get(c.vuln);
        if (!report) continue;
        c.title = report.title || "";
        c.status = report.status || "";
        c.extId = report.externalId || "";
      }
    }
  } catch {
    // enrichment is cosmetic — ids and dates are already complete
  }
  return cases;
}

async function pollNotifications(): Promise<void> {
  let token: string | null = null;
  try {
    token = await mintAccessToken();
  } catch {
    token = null;
  }
  if (!token) return; // no valid session — leave the badge as-is
  try {
    const url =
      `${NOTIFICATION_API}?$filter=${encodeURIComponent("read eq false")}&$count=true&$top=1`;
    const resp = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
    });
    if (!resp.ok) return;
    const raw: unknown = await resp.json();
    const count =
      raw && typeof raw === "object" && "@odata.count" in raw && typeof raw["@odata.count"] === "number"
        ? raw["@odata.count"]
        : 0;
    await chrome.action.setBadgeText({ text: count > 0 ? String(count) : "" });
    await chrome.action.setBadgeBackgroundColor({ color: "#a4262c" });
  } catch {
    // network hiccup — the next alarm retries
  }
}

async function startBadgePolling(): Promise<void> {
  await installOriginRules();
  chrome.alarms.create(POLL_ALARM, { periodInMinutes: 15 });
  void pollNotifications();
}

chrome.runtime.onStartup.addListener(() => void startBadgePolling());
chrome.runtime.onInstalled.addListener(() => void startBadgePolling());
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === POLL_ALARM) void pollNotifications();
});
// also poll once whenever the worker wakes (covers restarts without events)
void startBadgePolling();

// The toolbar icon cannot both show a popup and receive a double-click.
// default_popup is unset so onClicked fires. A second click inside the
// double-click window opens the portal already logged in (the new tab is
// seeded from the persisted MSAL snapshot). A single click still shows the
// new-message list.
const PORTAL_LIST_URL = MSRC_ORIGIN + "/report/vulnerability";
const DOUBLE_CLICK_MS = 600;
let pendingClick = 0;

function isPortalList(url: string | undefined): boolean {
  if (!url) return false;
  try {
    const parsed = new URL(url);
    const path = parsed.pathname.replace(/\/+$/, "");
    return parsed.origin === MSRC_ORIGIN && path === "/report/vulnerability";
  } catch {
    return false;
  }
}

async function openPortal(): Promise<void> {
  const tabs = await chrome.tabs.query({ url: MSRC_ORIGIN + "/*" });
  const existing = tabs.find((t) => t.id !== undefined && isPortalList(t.url));
  if (existing?.id !== undefined) {
    await chrome.tabs.update(existing.id, { active: true });
    if (existing.windowId !== undefined) {
      await chrome.windows.update(existing.windowId, { focused: true });
    }
    return;
  }
  await chrome.tabs.create({ url: PORTAL_LIST_URL, active: true });
}

async function showMessagePopup(): Promise<void> {
  await chrome.action.setPopup({ popup: "popup.html" });
  try {
    await chrome.action.openPopup();
  } catch {
    // leave the new-message list; do not turn a failed popup into the portal
    await chrome.action.setPopup({ popup: "" });
    await chrome.windows.create({
      url: chrome.runtime.getURL("popup.html"),
      type: "popup",
      width: 380,
      height: 520,
    });
    return;
  }
  // must be empty again or the next icon click is eaten by the popup
  await chrome.action.setPopup({ popup: "" });
}

chrome.action.onClicked.addListener(() => {
  if (pendingClick !== 0) {
    pendingClick = 0;
    void openPortal();
    return;
  }
  const token = Date.now();
  pendingClick = token;
  setTimeout(() => {
    if (pendingClick !== token) return;
    pendingClick = 0;
    void showMessagePopup();
  }, DOUBLE_CLICK_MS);
});

export {};
