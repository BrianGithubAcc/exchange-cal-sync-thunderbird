# Exchange Calendar Sync for Thunderbird

A Thunderbird extension that syncs calendars from Microsoft Exchange to Thunderbird using **EWS (Exchange Web Services)**.

## Features

- **Two-way sync** — events created, modified, or deleted in either Thunderbird or Exchange are synchronized
- **Multiple accounts** — sync calendars from several Exchange accounts simultaneously
- **OAuth 2.0 (device code flow)** — sign in with your Microsoft account; no Azure AD app registration required
- **Basic Auth** — alternative for on-premises Exchange servers that support it
- **Autodiscover** — automatically finds your Exchange EWS endpoint
- **Periodic sync** — configurable interval (default: 5 minutes)
- **Incremental sync** — initial full download, then efficient change-only sync via `SyncFolderItems`
- **Meeting invitations** — detects incoming iCalendar (iMIP) invites and syncs them
- **Meeting responses** — accept, tentatively accept, or decline meetings directly from Thunderbird
- **Full meeting details** — syncs attendees, organizer, HTML body (Teams links, etc.)

## Compatibility

- **Thunderbird** 128.0 – 150.*
- **Exchange** Microsoft 365 (Exchange Online), Exchange Server 2013/2016/2019

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

### Basic Auth

> **Note:** Microsoft 365 has disabled Basic Auth for most tenants. Use OAuth 2.0 instead.

1. Click the **"or use Basic Auth"** toggle in the options page
2. Enter your email, username (if different), and password
3. Click **Add Account**

## Architecture

```
manifest.json                → Extension manifest (Manifest V2)
background.html              → Background page (script loader)
background.js                → Lifecycle management, OAuth, message handling
ews/
  xml-builder.js             → EWS SOAP XML request construction
  autodiscover.js            → Autodiscover protocol implementation
  client.js                  → EWS HTTP client with auth + token refresh
sync/
  engine.js                  → Two-way sync engine with state tracking
  imip.js                    → iCalendar (iMIP) meeting invite processing
api/CalendarAPI/
  schema.json                → Experiment API schema
  implementation.js          → Bridge to Thunderbird's internal calendar API
options/
  options.html / options.js  → Settings UI (multi-account)
meetingbar/
  popup.html / popup.js      → Meeting response toolbar popup
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
- **Microsoft 365**: Basic Auth is likely disabled. Use OAuth 2.0.
- **On-premises Exchange**: Try `DOMAIN\username` or `username@domain.com` as the username.

### "Autodiscover failed"
- Enter the EWS URL manually in the options page.
- Microsoft 365: `https://outlook.office365.com/EWS/Exchange.asmx`

### Events not appearing
- Ensure the Exchange calendar is visible in Thunderbird's calendar sidebar.
- Try **Full Re-sync** from the options page.

### Attendees/meeting details missing
- Do a **Full Re-sync** — this fetches complete item details from Exchange.

## License

[MIT](LICENSE)
