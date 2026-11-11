// MSRC Helper — MAIN-world API proxy.
//
// Runs in the page's own JS world (manifest: "world": "MAIN") so its XHR
// carries the page origin. The portal APIs reject chrome-extension:// origins
// (CORS on /message, 403 on /portal/v2.0 REST), and even an isolated-world
// XHR is attributed to the extension — only a request issued from the page
// world behaves exactly like the SPA's own traffic.
//
// Protocol: the isolated content script dispatches `msrc-helper-apiget` with
// { url, nonce, body? }. Without `body` the request is a GET; with a string
// `body` it becomes a POST with Content-Type: application/json (used for
// GraphQL queries). The Bearer token is read from this world's own
// sessionStorage — never passed through the event, so the DOM event bus is
// not a token broadcast channel. This script replies with a
// `msrc-helper-apiget-reply-<nonce>` event carrying { status, body } or
// { error }.

function proxyAccessToken(): string | null {
  for (let i = 0; i < sessionStorage.length; i++) {
    const key = sessionStorage.key(i);
    if (!key || !key.includes("|accesstoken|")) continue;
    try {
      const parsed: unknown = JSON.parse(sessionStorage.getItem(key) || "");
      if (parsed && typeof parsed === "object" && "secret" in parsed) {
        const secret = (parsed as { secret: unknown }).secret;
        if (typeof secret === "string" && secret) return secret;
      }
    } catch {
      // malformed entry — keep scanning
    }
  }
  return null;
}

window.addEventListener("msrc-helper-apiget", (ev: Event) => {
  const detail = (ev as CustomEvent).detail as {
    url?: unknown;
    nonce?: unknown;
    body?: unknown;
  };
  const { url, nonce, body } = detail || {};
  if (typeof url !== "string" || typeof nonce !== "string") {
    return; // malformed request — ignore
  }
  if (!url.startsWith("https://api.msrc.microsoft.com/")) {
    return; // the proxy only ever talks to the portal API
  }
  const token = proxyAccessToken();
  if (!token) {
    window.dispatchEvent(
      new CustomEvent(`msrc-helper-apiget-reply-${nonce}`, { detail: { error: "no token" } })
    );
    return;
  }

  const xhr = new XMLHttpRequest();
  if (typeof body === "string") {
    xhr.open("POST", url);
    xhr.setRequestHeader("Content-Type", "application/json");
  } else {
    xhr.open("GET", url);
  }
  xhr.setRequestHeader("Authorization", `Bearer ${token}`);
  xhr.setRequestHeader("Accept", "application/json");

  const reply = (payload: unknown) =>
    window.dispatchEvent(
      new CustomEvent(`msrc-helper-apiget-reply-${nonce}`, { detail: payload })
    );

  xhr.onload = () => {
    try {
      reply({ status: xhr.status, body: JSON.parse(xhr.responseText) });
    } catch {
      reply({ status: xhr.status, error: "bad json" });
    }
  };
  xhr.onerror = () => reply({ error: "xhr" });
  xhr.send(typeof body === "string" ? body : undefined);
});
