# Exchange Calendar Sync for Thunderbird

A Thunderbird extension that syncs calendars from Microsoft Exchange to Thunderbird's
built-in Calendar using **EWS (Exchange Web Services)**. This checkout is configured
for Microsoft 365 OAuth using the Thunderbird public client and the EWS scope tested
by `scripts/test_ews_calendar.py`.

## Features

- **Two-way sync** — events created, modified, or deleted in either Thunderbird or Exchange are synchronized
- **Multiple accounts** — sync calendars from several Exchange accounts simultaneously
- **OAuth 2.0 (device code flow)** — sign in with your Microsoft account; no new Graph app registration
- **Secure token storage** — refresh tokens are stored in Thunderbird's Login Manager; access tokens stay in memory
- **Endpoint pinning** — bearer tokens are sent only to Microsoft's fixed EWS endpoint
- **Periodic sync** — configurable interval (default: 5 minutes)
- **Incremental sync** — initial full download, then efficient change-only sync via `SyncFolderItems`
- **Calendar events** — syncs titles, descriptions, locations, times, all-day state, and basic reminders

## Compatibility

- **Thunderbird** 128.0 – 153.* (ESR recommended)
- **Exchange** Microsoft 365 (Exchange Online) only

## Installation

### From XPI

1. Build the extension: `node build-xpi.js`
2. In Thunderbird: **Menu → Add-ons and Themes → ⚙ → Install Add-on From File**
3. Select `exchange-calendar-sync.xpi`

### For development

1. In Thunderbird: **Menu → Add-ons and Themes → ⚙ → Debug Add-ons**
2. Click **Load Temporary Add-on**
3. Select `manifest.json` from this directory

## Setup

### OAuth 2.0 (recommended)

1. Open extension options (**Add-ons → Exchange Calendar Sync → Options**)
2. Click **Sign in with Microsoft**
3. Go to the URL shown and enter the device code
4. After sign-in completes, enter your email and click **Add Account**
5. Sync starts automatically

The Microsoft 365 path uses OAuth only. No password or Basic Auth mode is accepted.

## Architecture

```
manifest.json                → Extension manifest (Manifest V2)
background.html              → Background page (script loader)
background.js                → Lifecycle management, OAuth, message handling
ews/
  xml-builder.js             → EWS SOAP XML request construction
  client.js                  → EWS HTTP client, endpoint policy, auth + token refresh
sync/
  engine.js                  → Two-way sync engine with state tracking
api/CalendarAPI/
  schema.json                → Experiment API schema
  implementation.js          → Bridge to Thunderbird's internal calendar API
api/SecureStorage/
  schema.json                → Login Manager storage API schema
  implementation.js          → Privileged refresh-token storage
options/
  options.html / options.js  → Settings UI (multi-account)
icons/                       → Extension icons (SVG)
build-xpi.js                 → Build script (Node.js, no dependencies)
```

## Building

```bash
node build-xpi.js
```

Produces `exchange-calendar-sync.xpi`. Requires Node.js (no npm dependencies).

## Troubleshooting

### "Authentication failed (HTTP 401)"
- Sign in again with Microsoft OAuth. Access tokens are memory-only and refresh tokens are stored in Thunderbird's Login Manager.

### Events not appearing
- Ensure the Exchange calendar is visible in Thunderbird's calendar sidebar.
- Try **Full Re-sync** from the options page.


## License

[MIT](LICENSE)
