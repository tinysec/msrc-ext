// MSRC Helper — content script (isolated world, document_start).
//
// Part 1 (every portal page): session seeding. Copies the extension-cached
// msal.* snapshot into this tab's sessionStorage before the SPA bundle runs,
// so a newly opened tab boots logged-in instead of bouncing to the B2C
// login. Also publishes this tab's (possibly refreshed) tokens back.
//
// Part 2 (report list page): a "Waiting(N)" pivot chip styled after the
// portal's own status pivots, showing the global count of cases whose latest
// thread message is the researcher's own. Clicking it renders those cases
// into the portal's own grid as native-looking rows. The grid also gains
// "External ID" and "CVE" columns, and column widths are normalized so no
// horizontal scrollbar appears. "/" focuses the search box (the portal has
// no such binding of its own).
//
// API proxy: relays background API calls to the MAIN-world proxy script
// (proxy.js), whose XHR carries the page origin — the portal APIs reject
// extension-origin requests.

const MSAL_PREFIX = "msal";
const MSRC_ORIGIN = "https://msrc.microsoft.com";
const SNAPSHOT_KEY = "msalSnapshot";
const PIVOT_CHIP_ID = "msrc-helper-waiting-pivot";

type Snapshot = Record<string, string>;

/* ---------------- Part 1: session seeding ---------------- */

function readMsalEntries(): Snapshot {
  const out: Snapshot = {};
  for (let i = 0; i < sessionStorage.length; i++) {
    const key = sessionStorage.key(i);
    if (key && key.startsWith(MSAL_PREFIX)) {
      const value = sessionStorage.getItem(key);
      if (value !== null) out[key] = value;
    }
  }
  return out;
}

function seedIfEmpty(snap: Snapshot | undefined): void {
  if (!snap || Object.keys(snap).length === 0) return;
  if (Object.keys(readMsalEntries()).length > 0) return;
  for (const [key, value] of Object.entries(snap)) {
    try {
      sessionStorage.setItem(key, value);
    } catch {
      // storage full or disabled — best effort
    }
  }
}

// Observability for the seeding path, exposed through the debug channel
// below. Kept tiny: it diagnosed every real-world failure so far.
const seedDebug = { attempts: 0, storageRead: "n/a", msgResult: "n/a", seeded: 0 };

async function bootstrapSeed(): Promise<void> {
  if (Object.keys(readMsalEntries()).length > 0) return; // already logged in
  for (let attempt = 0; attempt < 4; attempt++) {
    seedDebug.attempts = attempt + 1;
    let snap = await readCachedSnapshot();
    if (!snap || Object.keys(snap).length === 0) {
      snap = await requestSnapshot();
    }
    if (snap && Object.keys(snap).length > 0) {
      seedIfEmpty(snap);
      seedDebug.seeded = Object.keys(readMsalEntries()).length;
      return;
    }
    // the source tab may not have published yet — wait and retry
    const { promise: retryTick, resolve: tick } = Promise.withResolvers<void>();
    window.setTimeout(tick, 300);
    await retryTick;
  }
}

// Fast path: read the snapshot straight from session storage. Throws in a
// content script until the background has widened the access level.
async function readCachedSnapshot(): Promise<Snapshot | undefined> {
  try {
    const stored = await chrome.storage.session.get(SNAPSHOT_KEY);
    const snap = stored[SNAPSHOT_KEY] as Snapshot | undefined;
    seedDebug.storageRead = snap && Object.keys(snap).length > 0 ? "hit" : "empty";
    return snap;
  } catch {
    seedDebug.storageRead = "denied";
    return undefined;
  }
}

// Slow path: ask the background, which can fall back to live portal tabs.
async function requestSnapshot(): Promise<Snapshot | undefined> {
  try {
    const snap = (await chrome.runtime.sendMessage({
      type: "getSnapshot",
    })) as Snapshot | undefined;
    seedDebug.msgResult = snap && Object.keys(snap).length > 0 ? "hit" : "empty";
    return snap;
  } catch (e) {
    seedDebug.msgResult = "error:" + String(e).slice(0, 60);
    return undefined;
  }
}


function pushSnapshot(): void {
  try {
    void chrome.runtime
      .sendMessage({ type: "pushSnapshot", entries: readMsalEntries() })
      .catch(() => undefined);
  } catch {
    // extension context invalidated mid-reload
  }
}

/* ---------------- Part 2: "Waiting" pivot chip + grid view ---------------- */

interface WaitingCase {
  vuln: string;
  title: string;
  status: string;
  waitedSince: string;
}

const waitingState = {
  cases: [] as WaitingCase[], // every case waiting on MSRC, globally
  active: false, // is the Waiting grid view currently on
  inflight: false,
};

const chipDebug = { syncs: 0, lastError: "n/a", lastFetchSize: -1, appended: false };

// The label shows "…" while the very first computation is still running —
// a confident "Waiting(0)" during that window reads as a wrong answer.
function chipLabel(count: number | null): string {
  return count === null ? "Waiting(…)" : `Waiting(${count})`;
}

// The portal renders its status pivots as buttons labeled "Name(count)".
function findPivotButtons(): HTMLButtonElement[] {
  const pivots = Array.from(document.querySelectorAll("button")).filter((b) =>
    /^\S+\(\d+\)$/.test((b.innerText || "").trim())
  );
  return pivots.length >= 3 ? pivots : [];
}

function selectedPivotClass(pivots: HTMLButtonElement[]): string | null {
  const selected = pivots.find(
    (p) => p.getAttribute("aria-selected") === "true" || /selected/i.test(p.className)
  );
  return selected ? selected.className : null;
}


// Clicking the chip renders the waiting cases INTO the portal's own grid —
// cloned native rows, real columns, real links — instead of a side panel.
// The current page's real rows are hidden while the view is active.

const WAITING_ROW_ATTR = "data-msrc-helper-row";

interface ReportMeta {
  ext: string;
  title: string;
  status: string;
  ticket: string;
  impact: string;
}

let metaCacheKey = "";
let cveCache: Record<string, string> = {};
let metaCache: Record<string, ReportMeta> = {};

async function ensureMeta(ids: string[]): Promise<Record<string, ReportMeta>> {
  if (!extensionAlive()) return {};
  const key = ids.slice().sort().join(",");
  const resp = (await chrome.runtime.sendMessage({ type: "getReportMeta", ids })) as {
    records?: Record<string, ReportMeta>;
    cves?: Record<string, string>;
    error?: string;
  };
  if (resp?.error || !resp?.records) return {};
  metaCacheKey = key;
  metaCache = resp.records;
  cveCache = resp.cves || {};
  return metaCache;
}

function humanizeEnum(value: string): string {
  if (!value) return "";
  return value
    .toLowerCase()
    .split(/[_\s]+/)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

function cellIndexByHeader(): Map<string, number> {
  const map = new Map<string, number>();
  const headers = Array.from(document.querySelectorAll('[role="columnheader"], th'));
  headers.forEach((h, i) => {
    const name = (h.textContent || "").trim().toLowerCase();
    if (name && !map.has(name)) map.set(name, i);
  });
  return map;
}

function fillWaitingRow(
  row: HTMLElement,
  c: WaitingCase,
  meta: ReportMeta | undefined,
  cve: string,
  indexes: Map<string, number>
): void {
  const cells = Array.from(
    row.querySelectorAll('[role="gridcell"], td, [class*="-cell"]')
  ) as HTMLElement[];
  const set = (headerName: string, text: string) => {
    const idx = indexes.get(headerName);
    if (idx !== undefined && cells[idx]) cells[idx].textContent = text;
  };
  // Title cell carries the case link.
  const link = row.querySelector<HTMLAnchorElement>('a[href*="/VULN-"]');
  if (link) {
    link.setAttribute("href", `/report/vulnerability/${c.vuln}`);
    link.textContent = meta?.title || c.title || c.vuln;
  }
  set("status", humanizeEnum(meta?.status || c.status));
  set("case number", meta?.ticket || "—");
  set("security impact", humanizeEnum(meta?.impact || ""));
  set("bounty", "—");
  set("last modified", c.waitedSince ? new Date(c.waitedSince).toLocaleDateString() : "—");
  set("message info", "");
  set("status reason", "");
  const extCell = row.querySelector(`[${EXT_CELL_ATTR}="ext-cell"]`);
  if (extCell) extCell.textContent = meta?.ext || "—";
  const cveCell = row.querySelector(`[${EXT_CELL_ATTR}="cve-cell"]`);
  if (cveCell) cveCell.textContent = cve || "—";
}

async function applyWaitingView(on: boolean): Promise<void> {
  const grid = document.querySelector('[role="grid"]');
  const rowsParent = (grid?.querySelector('[role="rowgroup"], tbody') || grid) as HTMLElement | null;
  if (!rowsParent) return;

  // Remove any previously injected rows and restore hidden ones first, so
  // the operation below always starts from the portal's own DOM.
  for (const injected of document.querySelectorAll(`[${WAITING_ROW_ATTR}]`)) {
    injected.remove();
  }
  const realRows = Array.from(document.querySelectorAll(".ms-DetailsRow")).filter(
    (r) => !r.hasAttribute(WAITING_ROW_ATTR)
  ) as HTMLElement[];
  for (const row of realRows) row.style.display = "";
  if (rowsParent.dataset.msrcHelperHeight !== undefined) {
    rowsParent.style.height = rowsParent.dataset.msrcHelperHeight;
    delete rowsParent.dataset.msrcHelperHeight;
  }
  if (!on || waitingState.cases.length === 0) return;

  const template = realRows[0]?.cloneNode(true) as HTMLElement | undefined;
  if (!template) return; // no native row to mirror — leave the grid alone
  // Fluent's DetailsList positions rows absolutely at i*rowHeight inside a
  // canvas-sized rowgroup; mirror that layout or the rows land in normal
  // flow BELOW the reserved canvas (the blank gap).
  const rowHeight = realRows[0].offsetHeight || 42;
  const meta = await ensureMeta(waitingState.cases.map((c) => c.vuln));
  const cves = cveCache;
  const indexes = cellIndexByHeader();
  const sorted = [...waitingState.cases].sort(
    (a, b) => Date.parse(a.waitedSince) - Date.parse(b.waitedSince)
  );
  for (const row of realRows) row.style.display = "none";
  rowsParent.dataset.msrcHelperHeight = rowsParent.style.height || "";
  rowsParent.style.height = `${sorted.length * rowHeight}px`;
  sorted.forEach((c, i) => {
    const row = template.cloneNode(true) as HTMLElement;
    row.setAttribute(WAITING_ROW_ATTR, c.vuln);
    row.style.position = "absolute";
    row.style.top = `${i * rowHeight}px`;
    row.style.left = "0";
    row.style.width = "100%";
    row.style.transform = "none";
    row.style.display = "";
    fillWaitingRow(row, c, meta[c.vuln], cves[c.vuln] || "", indexes);
    rowsParent.appendChild(row);
  });
}

/* ---------------- "External ID" column ---------------- */

const EXT_CELL_ATTR = "data-msrc-helper-ext";
let extIdsKey = "";

// The list hides the short MSRC case id (externalId, e.g. "197") that vendor
// correspondence actually references. Append an "External ID" column to the
// grid and hide the redundant "Report ID" column to make room. Data rides
// the page-origin proxy; when the report API is down the grid stays as-is.
async function syncExternalIds(): Promise<void> {
  if (!extensionAlive()) return;
  const ids = Array.from(
    new Set(
      Array.from(
        document.querySelectorAll<HTMLAnchorElement>('a[href^="/report/vulnerability/VULN-"]')
      ).map((a) => (a.getAttribute("href") || "").split("/").pop() || "")
    )
  ).filter(Boolean);
  if (ids.length === 0) return;
  const key = ids.join(",");
  if (key === extIdsKey) return;

  const records = await ensureMeta(ids);
  if (Object.keys(records).length === 0) return; // no data — retry on the next render
  extIdsKey = key; // only mark done once the columns actually filled
  const cves = cveCache;

  // Header: hide "Report ID", append "External ID" and "CVE".
  const headers = Array.from(document.querySelectorAll('[role="columnheader"], th'));
  const reportIdHeader = headers.find(
    (h): h is HTMLElement => /^Report ID$/i.test((h.textContent || "").trim())
  );
  if (reportIdHeader) reportIdHeader.style.display = "none";
  const headerRow = headers[0]?.parentElement;
  if (headerRow) {
    const donorHeader = (reportIdHeader || headers[headers.length - 1]) as HTMLElement | undefined;
    for (const [marker, label] of [
      ["ext-header", "External ID"],
      ["cve-header", "CVE"],
    ] as const) {
      if (headerRow.querySelector(`[${EXT_CELL_ATTR}="${marker}"]`)) continue;
      const th = donorHeader?.cloneNode(true) as HTMLElement;
      th.setAttribute(EXT_CELL_ATTR, marker);
      th.style.display = "";
      th.textContent = label;
      headerRow.appendChild(th);
    }
  }

  // Rows: hide the Report ID cell (same index as its header), append the
  // external id and any CVEs parsed out of the thread messages.
  const reportIdIndex = reportIdHeader ? headers.indexOf(reportIdHeader) : -1;
  for (const row of document.querySelectorAll(".ms-DetailsRow, [role='row']")) {
    if (row.hasAttribute(WAITING_ROW_ATTR)) continue; // injected rows are pre-filled
    const link = row.querySelector<HTMLAnchorElement>('a[href*="/VULN-"]');
    if (!link) continue;
    const id = (link.getAttribute("href") || "").split("/").pop() || "";
    const cells = Array.from(row.querySelectorAll('[role="gridcell"], td, [class*="-cell"]'));
    if (reportIdIndex >= 0 && cells[reportIdIndex]) {
      (cells[reportIdIndex] as HTMLElement).style.display = "none";
    }
    const values: Array<[string, string]> = [
      ["ext-cell", records[id]?.ext || "—"],
      ["cve-cell", cves[id] || "—"],
    ];
    for (const [marker, text] of values) {
      let cell = row.querySelector(`[${EXT_CELL_ATTR}="${marker}"]`) as HTMLElement | null;
      if (!cell) {
        cell = document.createElement("div");
        const donor = cells[cells.length - 1] as HTMLElement | undefined;
        if (donor) cell.className = donor.className;
        cell.setAttribute(EXT_CELL_ATTR, marker);
        cell.style.cssText = "min-width:80px;display:flex;align-items:center;";
        row.appendChild(cell);
      }
      cell.textContent = text;
    }
  }
  fitColumns();
}

// The portal hard-codes wasteful column tracks (Bounty: 220px for a single
// character; both date columns: 200px for 16-character dates), which pushed
// the added columns into a horizontal scrollbar and left Title truncated at
// 320px. Deterministic fix: fixed content-driven widths for every known
// column, Title absorbs everything left over.
const COLUMN_TARGETS: Record<string, number> = {
  "message info": 44,
  status: 130,
  "status reason": 44,
  bounty: 70,
  "created on": 150,
  "last modified": 150,
  "case number": 90,
  "security impact": 160,
  "reported products": 110,
  feedback: 44,
  "delete draft": 44,
  "links to individual reports": 44,
  "external id": 90,
  cve: 150,
};

function setColumnWidth(headers: HTMLElement[], rows: HTMLElement[], index: number, px: number): void {
  const value = `${px}px`;
  headers[index].style.width = value;
  headers[index].style.minWidth = value;
  for (const row of rows) {
    const cell = row.querySelectorAll('[role="gridcell"], td, [class*="-cell"]')[index] as
      | HTMLElement
      | undefined;
    if (cell) {
      cell.style.width = value;
      cell.style.minWidth = value;
    }
  }
}

function fitColumns(): void {
  const grid = document.querySelector('[role="grid"]');
  const scroller = grid?.parentElement as HTMLElement | null | undefined;
  if (!grid || !scroller) return;
  const headers = Array.from(grid.querySelectorAll('[role="columnheader"], th')) as HTMLElement[];
  const rows = Array.from(grid.querySelectorAll(".ms-DetailsRow")) as HTMLElement[];
  if (headers.length === 0 || rows.length === 0) return;

  let fixedTotal = 0;
  let titleIndex = -1;
  headers.forEach((h, index) => {
    if (h.style.display === "none") return;
    const raw = (h.textContent || "").trim().toLowerCase();
    if (/title|description/.test(raw)) {
      titleIndex = index; // sized last from the leftover space
      return;
    }
    // Header text carries Fluent icon glyphs and may be truncated ("Links
    // to individu…"), so strip non-letters and match by prefix in both
    // directions instead of exact keys.
    const name = raw.replace(/[^a-z ]/g, " ").replace(/\s+/g, " ").trim();
    const targetKey = Object.keys(COLUMN_TARGETS).find(
      (k) => k === name || (name.length >= 6 && (k.startsWith(name) || name.startsWith(k)))
    );
    const target = targetKey !== undefined ? COLUMN_TARGETS[targetKey] : undefined;
    if (target !== undefined) {
      setColumnWidth(headers, rows, index, target);
      fixedTotal += target;
    } else {
      fixedTotal += h.offsetWidth; // unknown column — keep the portal's width
    }
  });
  if (titleIndex >= 0) {
    setColumnWidth(headers, rows, titleIndex, Math.max(280, scroller.clientWidth - fixedTotal));
  }
}

// Clone a native pivot button so the chip matches the portal's styling for
// free, then swap in our label and behavior.
function buildPivotChip(
  template: HTMLButtonElement,
  count: number | null,
  selectedClass: string | null
): HTMLButtonElement {
  const chip = template.cloneNode(true) as HTMLButtonElement;
  chip.removeAttribute("id");
  chip.id = PIVOT_CHIP_ID;
  chip.setAttribute("aria-selected", "false");
  chip.textContent = chipLabel(count);
  chip.title =
    "Cases where your message is the latest one in the thread (all cases, not just this view)";
  chip.addEventListener("click", () => {
    waitingState.active = !waitingState.active;
    chip.setAttribute("aria-selected", String(waitingState.active));
    // mirror the native selected look while pressed
    if (waitingState.active && selectedClass) chip.className = selectedClass;
    else chip.className = template.className;
    void applyWaitingView(waitingState.active);
  });
  return chip;
}

// Create the chip if missing, or update its count in place. Re-resolves the
// anchor on every call: the SPA re-renders the pivot bar freely, and a
// template captured before a re-render is detached from the document.
function ensureChip(count: number | null, fallbackPivots: HTMLButtonElement[]): HTMLButtonElement | null {
  const pivots = findPivotButtons();
  const anchorList = pivots.length > 0 ? pivots : fallbackPivots;
  const template = anchorList[anchorList.length - 1];

  const existing = document.getElementById(PIVOT_CHIP_ID) as HTMLButtonElement | null;
  if (existing) {
    existing.textContent = chipLabel(count);
    chipDebug.appended = true;
    return existing;
  }
  template.parentElement?.appendChild(buildPivotChip(template, count, selectedPivotClass(anchorList)));
  chipDebug.appended = !!document.getElementById(PIVOT_CHIP_ID);
  if (!chipDebug.appended) {
    // anchor was detached mid-render — retry shortly with fresh resolution
    window.setTimeout(() => void syncPivotChip(), 600);
  }
  return document.getElementById(PIVOT_CHIP_ID) as HTMLButtonElement | null;
}

// After an extension reload (↻), content scripts in already-open tabs keep
// running with a severed message port — "Extension context invalidated".
// chrome.runtime.id turns undefined in that state; detect it once and shut
// the machinery down silently instead of spamming console errors until the
// tab is refreshed.
let extensionDead = false;

function extensionAlive(): boolean {
  if (extensionDead) return false;
  if (chrome.runtime && chrome.runtime.id) return true;
  extensionDead = true;
  return false;
}

// Renders the chip instantly from the background cache, then refreshes the
// count in place when the cached data is stale — no pop-in delay, no flicker.
async function syncPivotChip(): Promise<void> {
  if (waitingState.inflight || !extensionAlive()) return;
  const pivots = findPivotButtons();
  if (pivots.length === 0) return;
  waitingState.inflight = true;
  chipDebug.syncs++;
  try {
    // Phase 1: cached render. No cache yet (fresh service worker, first run
    // after install) → show the pending label instead of a wrong zero.
    const cached = (await chrome.runtime.sendMessage({ type: "getAwaitingCached" })) as {
      cases?: WaitingCase[];
      ageMs?: number;
    } | undefined;
    if (cached?.cases && cached.cases.length > 0) {
      waitingState.cases = cached.cases;
      chipDebug.lastFetchSize = cached.cases.length;
      ensureChip(cached.cases.length, pivots);
      if ((cached.ageMs ?? Number.POSITIVE_INFINITY) < 60_000) {
        return; // fresh enough — skip the network round trip
      }
    } else {
      ensureChip(null, pivots);
    }

    // Phase 2: freshness-checked full refresh.
    const resp = (await chrome.runtime.sendMessage({ type: "getAwaiting" })) as {
      cases?: WaitingCase[];
      error?: string;
    };
    if (resp?.error || !resp?.cases) {
      throw new Error(resp?.error || "no awaiting data");
    }
    waitingState.cases = resp.cases;
    chipDebug.lastFetchSize = resp.cases.length;
    const chip = ensureChip(resp.cases.length, pivots);
    if (waitingState.active && chip) {
      chip.setAttribute("aria-selected", "true");
      void applyWaitingView(true); // re-render injected rows with fresh data
    }
  } catch (e) {
    chipDebug.lastError = String(e);
    if (!extensionAlive()) return; // orphaned by an extension reload — stay quiet
    console.warn("MSRC helper: awaiting filter unavailable", e);
  } finally {
    waitingState.inflight = false;
  }
}

/* ---------------- API proxy relay ---------------- */

// The background asks for a URL; the MAIN-world proxy performs the XHR with
// the page origin and reads the Bearer token from sessionStorage itself —
// the token never rides the DOM event bus. DOM CustomEvents cross the world
// boundary; the nonce matches request to reply.
function handleApiGet(
  url: unknown,
  body: unknown,
  sendResponse: (r: unknown) => void
): boolean {
  if (typeof url !== "string" || !url.startsWith("https://api.msrc.microsoft.com/")) {
    sendResponse({ error: "url not allowed" });
    return false;
  }
  const payload =
    typeof body === "string" ? { body } : {};
  const nonce = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const replyEvent = `msrc-helper-apiget-reply-${nonce}`;
  const { promise, resolve } = Promise.withResolvers<void>();
  const timer = window.setTimeout(() => {
    window.removeEventListener(replyEvent, onReply);
    sendResponse({ error: "proxy timeout" });
    resolve();
  }, 20000);
  const onReply = (ev: Event) => {
    window.clearTimeout(timer);
    sendResponse((ev as CustomEvent).detail);
    resolve();
  };
  window.addEventListener(replyEvent, onReply, { once: true });
  window.dispatchEvent(
    new CustomEvent("msrc-helper-apiget", { detail: { url, nonce, ...payload } })
  );
  void promise;
  return true; // async sendResponse
}

/* ---------------- debug channel ---------------- */

// DOM events cross the page/isolated world boundary, which makes them a
// probe point reachable from the page console:
//   dispatchEvent(new CustomEvent("msrc-helper-debug"))
// then listen for "msrc-helper-debug-reply".
window.addEventListener("msrc-helper-debug", () => {
  window.dispatchEvent(
    new CustomEvent("msrc-helper-debug-reply", {
      detail: {
        seed: { ...seedDebug, msalKeys: Object.keys(readMsalEntries()).length },
        chip: {
          ...chipDebug,
          chipText: document.getElementById(PIVOT_CHIP_ID)?.textContent || null,
          inDom: !!document.getElementById(PIVOT_CHIP_ID),
          pivots: findPivotButtons().length,
          waitingCount: waitingState.cases.length,
          active: waitingState.active,
        },
      },
    })
  );
});

/* ---------------- list-page mounting ---------------- */

function mountListTools(): void {
  let timer = 0;
  const schedule = () => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => {
      void syncPivotChip().then(() => fitColumns());
      void syncExternalIds();
    }, 600);
  };
  // The SPA re-renders freely; every structural change re-syncs the chip.
  new MutationObserver(schedule).observe(document.body, {
    childList: true,
    subtree: true,
  });
  // Coming back to the tab is the natural moment to refresh the count.
  window.addEventListener("focus", schedule);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") schedule();
  });
  schedule();

  // The portal itself has no "/" binding (verified against its bundles).
  document.addEventListener(
    "keydown",
    (e) => {
      if (e.key !== "/" || e.ctrlKey || e.altKey || e.metaKey) return;
      const target = e.target as HTMLElement | null;
      if (
        target &&
        (target.tagName === "INPUT" ||
          target.tagName === "TEXTAREA" ||
          target.isContentEditable)
      ) {
        return;
      }
      const search = document.querySelector<HTMLInputElement>(
        'input[aria-label*="Search" i], input[placeholder*="Search" i]'
      );
      if (search) {
        e.preventDefault();
        search.focus();
      }
    },
    true
  );
}

/* ---------------- bootstrap ---------------- */

function start(): void {
  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg?.type === "querySnapshot") {
      sendResponse(readMsalEntries());
      return false;
    }
    if (msg?.type === "apiGet") {
      return handleApiGet(msg.url, msg.body, sendResponse);
    }
    return false;
  });

  void bootstrapSeed().finally(() => {
    // document_start: the SPA bundle loads and parses after this script, so
    // by DOMContentLoaded the (possibly seeded) tokens are current — publish.
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", () => pushSnapshot(), { once: true });
    } else {
      pushSnapshot();
    }
    window.addEventListener("pagehide", () => pushSnapshot());
    window.setInterval(() => pushSnapshot(), 5 * 60 * 1000);
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => mountListTools(), { once: true });
  } else {
    mountListTools();
  }
}

// Defense in depth: the manifest injects this script only into
// https://msrc.microsoft.com top-level frames. Never read, write, or relay
// anything on any other origin, even if injection rules change someday.
if (location.origin === MSRC_ORIGIN) {
  start();
}

export {};
