# MSRC Researcher Portal Extension

[![build](https://github.com/tinysec/msrc-ext/actions/workflows/ci.yaml/badge.svg)](https://github.com/tinysec/msrc-ext/actions/workflows/ci.yaml)
[![release](https://img.shields.io/github/v/release/tinysec/msrc-ext?include_prereleases&display_name=tag)](https://github.com/tinysec/msrc-ext/releases)

Chrome extension (MV3) for the [MSRC Researcher Portal](https://msrc.microsoft.com/report/vulnerability). It keeps a signed-in session across tabs and browser restarts, and surfaces new case messages from the toolbar.

## Features

- **Signed-in new tabs.** The portal stores MSAL tokens in per-tab `sessionStorage` and sets no auth cookie, so a new tab would otherwise bounce to B2C login. This extension copies the `msal.*` snapshot in at `document_start`. Middle-click, Ctrl+click, and pasted URLs all work.
- **Session survives browser exit.** The refresh token is stored in `chrome.storage.local`. After a restart, a new tab is seeded from that copy. An expired access token is fine: MSAL renews it from the refresh token in the same snapshot.
- **Unread badge.** A 15-minute alarm mints an access token and polls unread notifications. The count is shown on the toolbar icon. No portal tab has to be open.
- **New-message list.** A single click on the icon lists `NewMessageOnReport` cases, newest first, unread rows first. It is not the Waiting list.
- **Open the portal.** Double-click the icon within 600ms. An existing tab whose path is exactly `/report/vulnerability` is focused. Otherwise a new tab opens at that URL, already signed in. An open case page is not treated as the list.
- **Waiting(N) on the portal page.** A pivot chip counts cases whose latest thread message is your own. Clicking it renders those cases into the portal grid. Click again to restore the portal's own rows.
- **List columns.** External ID and CVE are added. Report ID is hidden. Column widths are normalized so the grid does not grow a horizontal scrollbar.
- **`/` focuses search.** The portal has no such binding.

## Usage

1. Download `msrc-researcher-portal-extension-v0.1.N.zip` from [Releases](https://github.com/tinysec/msrc-ext/releases), or build locally and use `dist/`.
2. Unzip it. In Chrome, open `chrome://extensions`, turn on Developer mode, and choose **Load unpacked**. Select the unzipped folder (or `dist/`).
3. Open the portal once and sign in. That publishes the MSAL snapshot the extension keeps.
4. After an update, click **Reload** on the extension card, then refresh every open portal tab. A content script from the previous build stays alive until the page reloads.

| Action | Result |
| --- | --- |
| Single click the icon | New-message cases. Click a case number to open it. |
| Double-click the icon | Opens `https://msrc.microsoft.com/report/vulnerability`, signed in. |
| Hover the icon | Unread count, when there is one. |

The toolbar title is `Double-click to open the portal`.

## Build

Releases are built on a GitHub-hosted `ubuntu-latest` runner. The version is `v0.1.{run_number}`, where `{run_number}` is `github.run_number`. A push to `master` publishes that zip. Pull requests build only.

```bash
npm ci
npm run build
```

Load `dist/` as an unpacked extension. A local build keeps the version in `manifest.json` (`0.1.0`). The runner rewrites that field before it packs the release zip.

## Privacy

- Requests go only to `api.msrc.microsoft.com` and `msrcweb.b2clogin.com`. No telemetry.
- The Bearer token is not put on the DOM event bus. The page-world proxy reads it from `sessionStorage`.
- The refresh token is the only credential stored on disk. It is removed when a portal tab that had a session reports logout. Message bodies are not stored.
- `declarativeNetRequest` rules rewrite `Origin` only for requests this extension starts.
