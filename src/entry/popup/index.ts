// MSRC Researcher Portal Extension popup: cases whose latest thread message
// came from MSRC — i.e. cases with a new message waiting for you, newest
// first. This is the opposite of the portal page's Waiting view.

const MSRC_ORIGIN = "https://msrc.microsoft.com";

interface NewMessageCase {
  vuln: string;
  extId: string;
  title: string;
  status: string;
  createdOn: string;
}

function when(iso: string): string {
  const t = Date.parse(iso);
  if (!t) return "";
  return new Date(t).toISOString().slice(0, 10);
}

async function render(): Promise<void> {
  const meta = document.getElementById("meta") as HTMLElement;
  const list = document.getElementById("list") as HTMLElement;
  meta.textContent = "Loading…";
  list.replaceChildren();
  const resp = (await chrome.runtime.sendMessage({ type: "getNewMessages" })) as {
    cases?: NewMessageCase[];
    error?: string;
  };
  if (resp?.error) {
    meta.textContent = resp.error;
    meta.className = "err";
    return;
  }
  const cases = resp?.cases || [];
  meta.textContent =
    cases.length === 0 ? "No new messages." : `${cases.length} case(s) with a new message`;
  meta.className = "meta";
  for (const c of cases) {
    const li = document.createElement("li");
    const a = document.createElement("a");
    a.href = "#";
    a.textContent = c.extId || c.title || "";
    a.addEventListener("click", (e) => {
      e.preventDefault();
      void chrome.tabs.create({
        url: `${MSRC_ORIGIN}/report/vulnerability/${c.vuln}`,
        active: true,
      });
    });
    const title = document.createElement("span");
    title.className = "title";
    title.textContent = c.title || "(no title)";
    const sub = document.createElement("span");
    sub.className = "sub";
    const bits = [c.status, when(c.createdOn)].filter(Boolean);
    sub.textContent = bits.join(" · ");
    li.append(a, title, sub);
    list.appendChild(li);
  }
}

document.addEventListener("DOMContentLoaded", () => {
  void render();
  document.getElementById("refresh")?.addEventListener("click", () => void render());
  document.getElementById("open")?.addEventListener("click", () => {
    void chrome.tabs.create({ url: `${MSRC_ORIGIN}/report/vulnerability`, active: true });
  });
});

export {};
