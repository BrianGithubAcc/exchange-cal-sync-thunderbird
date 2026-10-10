console.log("[ExCalSync] engine.js loading");
/**
 * Two-way sync engine for Exchange <-> Thunderbird calendar.
 * Tracks sync state, handles conflict resolution, and manages item mappings.
 */

// Depends on: ews/client.js (EwsClient), ews/xml-builder.js (xmlEscape)

// Sync interval default: 5 minutes
const DEFAULT_SYNC_INTERVAL_MINUTES = 5;

// How far back/forward to sync (days)
const DEFAULT_SYNC_RANGE_PAST_DAYS = 30;
const DEFAULT_SYNC_RANGE_FUTURE_DAYS = 90;

/**
 * Convert an EWS calendar item to ICS format.
 */
function ewsItemToIcs(ewsItem) {
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Exchange Calendar Sync//EN",
    "BEGIN:VEVENT",
  ];

  // Use an injective encoding of the Exchange Item ID. A short 32-bit hash
  // can collide and would merge two unrelated events.
  const uid = ewsItem.id
    ? stableEwsUid(ewsItem.id)
    : `ews-${Date.now()}@exchange-cal-sync`;
  lines.push(`UID:${uid}`);

  if (ewsItem.subject) {
    lines.push(`SUMMARY:${icsEscapeValue(ewsItem.subject)}`);
  }
  if (ewsItem.body) {
    if (ewsItem.bodyType === "HTML") {
      // Strip HTML tags for plain text DESCRIPTION
      const plainText = ewsItem.body
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<\/p>/gi, "\n")
        .replace(/<\/div>/gi, "\n")
        .replace(/<[^>]+>/g, "")
        .replace(/&nbsp;/gi, " ")
        .replace(/&amp;/gi, "&")
        .replace(/&lt;/gi, "<")
        .replace(/&gt;/gi, ">")
        .replace(/&quot;/gi, '"')
        .replace(/&#39;/gi, "'")
        .trim();
      if (plainText) {
        lines.push(`DESCRIPTION:${icsEscapeValue(plainText)}`);
      }
      // Preserve the full HTML body so Thunderbird can render clickable links
      // ICS long lines must be folded at 75 octets; use X-ALT-DESC for HTML
      const escapedHtml = icsEscapeValue(ewsItem.body);
      lines.push(`X-ALT-DESC;FMTTYPE=text/html:${escapedHtml}`);
    } else {
      lines.push(`DESCRIPTION:${icsEscapeValue(ewsItem.body)}`);
    }
  }
  if (ewsItem.location) {
    lines.push(`LOCATION:${icsEscapeValue(ewsItem.location)}`);
  }
  if (ewsItem.start) {
    lines.push(`DTSTART:${dateToIcsFormat(ewsItem.start, ewsItem.isAllDay)}`);
  }
  if (ewsItem.end) {
    lines.push(`DTEND:${dateToIcsFormat(ewsItem.end, ewsItem.isAllDay)}`);
  }
  if (ewsItem.isAllDay) {
    lines.push("X-MICROSOFT-CDO-ALLDAYEVENT:TRUE");
  }
  if (ewsItem.sensitivity && ewsItem.sensitivity !== "Normal") {
    const classMap = {
      Confidential: "CONFIDENTIAL",
      Private: "PRIVATE",
      Personal: "PRIVATE",
    };
    lines.push(`CLASS:${classMap[ewsItem.sensitivity] || "PUBLIC"}`);
  }
  if (ewsItem.importance) {
    const prioMap = { High: "1", Normal: "5", Low: "9" };
    lines.push(`PRIORITY:${prioMap[ewsItem.importance] || "5"}`);
  }
  if (ewsItem.categories && ewsItem.categories.length > 0) {
    lines.push(
      `CATEGORIES:${ewsItem.categories.map(icsEscapeValue).join(",")}`
    );
  }
  if (ewsItem.reminderSet && ewsItem.reminderMinutes > 0) {
    lines.push("BEGIN:VALARM");
    lines.push("ACTION:DISPLAY");
    lines.push("DESCRIPTION:Reminder");
    lines.push(`TRIGGER:-PT${ewsItem.reminderMinutes}M`);
    lines.push("END:VALARM");
  }
  if (ewsItem.organizer) {
    const orgCn = ewsItem.organizer.name ? `;CN=${icsEscapeValue(ewsItem.organizer.name)}` : "";
    const organizerEmail = safeIcsEmail(ewsItem.organizer.email);
    if (organizerEmail) {
      lines.push(`ORGANIZER${orgCn}:mailto:${organizerEmail}`);
    }
  }
  if (ewsItem.attendees) {
    for (const att of ewsItem.attendees) {
      const role =
        att.type === "optional" ? "OPT-PARTICIPANT" : "REQ-PARTICIPANT";
      const attendeeEmail = safeIcsEmail(att.email);
      if (attendeeEmail) {
        const cn = att.name ? `;CN=${icsEscapeParameter(att.name)}` : "";
        lines.push(`ATTENDEE;ROLE=${role}${cn}:mailto:${attendeeEmail}`);
      }
    }
  }
  if (ewsItem.lastModified) {
    lines.push(
      `LAST-MODIFIED:${dateToIcsFormat(ewsItem.lastModified, false)}`
    );
  }

  lines.push("END:VEVENT");
  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}

/**
 * Parse basic ICS fields from ICS data for pushing to Exchange.
 * Handles TZID parameters and various date formats from Thunderbird.
 */
function icsToEwsFields(icsData) {
  const fields = {};
  // Unfold continued lines (RFC 5545: line starts with space = continuation)
  const lines = icsData.replace(/\r\n /g, "").replace(/\r\n\t/g, "").split(/\r?\n/);

  for (const line of lines) {
    const colonIdx = line.indexOf(":");
    if (colonIdx === -1) continue;
    const keyPart = line.substring(0, colonIdx);
    const key = keyPart.split(";")[0];
    const value = line.substring(colonIdx + 1);

    switch (key) {
      case "SUMMARY":
        fields.subject = icsUnescapeValue(value);
        break;
      case "DESCRIPTION":
        fields.body = icsUnescapeValue(value);
        break;
      case "LOCATION":
        fields.location = icsUnescapeValue(value);
        break;
      case "DTSTART":
        fields.start = icsDateToIso(value);
        fields.isAllDay = value.length === 8;
        break;
      case "DTEND":
        fields.end = icsDateToIso(value);
        break;
      case "CLASS":
        fields.sensitivity = {
          PRIVATE: "Private",
          CONFIDENTIAL: "Confidential",
        }[value.toUpperCase()] || "Normal";
        break;
    }
  }

  const reminderMatch = icsData.match(/TRIGGER:-PT(\d+)M/i);
  if (reminderMatch) {
    fields.reminderSet = true;
    fields.reminderMinutes = Number.parseInt(reminderMatch[1], 10);
  }

  if (fields.isAllDay) {
    // EWS requires DateTime values even when IsAllDayEvent is true. Keep the
    // date at midnight UTC so the date does not shift across local timezones.
    if (fields.start && fields.start.length === 10) {
      fields.start += "T00:00:00Z";
    }
    if (fields.end && fields.end.length === 10) {
      fields.end += "T00:00:00Z";
    }
  }

  // If we have a start but no end, default end = start + 1 hour (or +1 day for all-day)
  if (fields.start && !fields.end) {
    const d = new Date(fields.start);
    if (!isNaN(d.getTime())) {
      if (fields.isAllDay) {
        d.setDate(d.getDate() + 1);
        fields.end = d.toISOString();
      } else {
        d.setHours(d.getHours() + 1);
        fields.end = d.toISOString();
      }
    }
  }

  return fields;
}

/**
 * Convert ISO date string to ICS format.
 */
function dateToIcsFormat(isoString, isAllDay) {
  if (isAllDay) {
    // For all-day events, extract date directly to avoid timezone shifting
    const dateMatch = isoString.match(/^(\d{4})-?(\d{2})-?(\d{2})/);
    if (dateMatch) {
      return dateMatch[1] + dateMatch[2] + dateMatch[3];
    }
  }
  const d = new Date(isoString);
  if (isNaN(d.getTime())) return isoString;

  return d.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}/, "");
}

/**
 * Convert ICS date to ISO format for EWS.
 */
function icsDateToIso(icsDate) {
  if (icsDate.length === 8) {
    // All-day: YYYYMMDD
    return `${icsDate.substring(0, 4)}-${icsDate.substring(4, 6)}-${icsDate.substring(6, 8)}`;
  }
  // DateTime: YYYYMMDDTHHMMSSZ or YYYYMMDDTHHMMSS (floating)
  const hasUtcSuffix = icsDate.endsWith("Z");
  const clean = icsDate.replace("Z", "");
  const formatted = `${clean.substring(0, 4)}-${clean.substring(4, 6)}-${clean.substring(6, 8)}T${clean.substring(9, 11)}:${clean.substring(11, 13)}:${clean.substring(13, 15)}`;
  if (hasUtcSuffix) {
    return formatted + "Z";
  }
  // Floating time (no Z suffix): convert from local time to UTC
  const d = new Date(formatted);
  if (!isNaN(d.getTime())) {
    return d.toISOString();
  }
  return formatted;
}

function icsEscapeValue(str) {
  if (!str) return "";
  return str
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\n/g, "\\n")
    .replace(/\r/g, "");
}

function icsEscapeParameter(str) {
  if (!str) return "";
  return String(str).replace(/[\\;,\r\n]/g, function (ch) {
    return ch === "\\" ? "\\\\" : "\\" + ch;
  });
}

function safeIcsEmail(value) {
  if (typeof value !== "string" || value.length > 320 ||
      /[\r\n\x00-\x1f\x7f]/.test(value)) return null;
  const email = value.trim();
  if (!/^[^\s@<>]+@[^\s@<>]+$/.test(email)) return null;
  return email;
}

function icsUnescapeValue(str) {
  if (!str) return "";
  return str
    .replace(/\\n/g, "\n")
    .replace(/\\,/g, ",")
    .replace(/\\;/g, ";")
    .replace(/\\\\/g, "\\");
}

function stableEwsUid(id) {
  return `ews-${encodeURIComponent(String(id))}@exchange-cal-sync`;
}

function isEwsConflictError(error) {
  return /changekey|conflict|itemnotfound/i.test(String(error && error.message));
}

/**
 * The Sync Engine manages the sync state and orchestrates sync operations.
 */
class SyncEngine {
  constructor() {
    this.ewsClient = null;
    this.calendarId = null;
    this.syncState = null;
    this.itemMap = {}; // Maps: ewsItemId -> { localItemId, changeKey, lastModified }
    this.reverseMap = {}; // Maps: localItemId -> ewsItemId
    this._syncing = false;
    this._localChangesQueue = []; // Queue of local changes to push upstream
    this.conflicts = {}; // local item id -> conflict details
    this.lastError = null;
    this._onLocalChange = null;
    this._initialized = false;
    this._suppressLocalEvents = false; // Suppress observer during sync-from-remote
  }

  /**
   * Initialize the sync engine.
   */
  async init(config) {
    console.log("[SyncEngine] Initializing...");

    // State (syncState, itemMap, reverseMap, calendarId) is pre-set by caller
    console.log("[SyncEngine] Pre-set state: calendarId=", this.calendarId, "items=", Object.keys(this.itemMap).length);

    // Create EWS client
    console.log("[SyncEngine] Creating EwsClient, authType=", config.credentials.type, "ewsUrl=", config.ewsUrl);
    this.ewsClient = new EwsClient({
      email: config.email,
      ewsUrl: config.ewsUrl || null,
      credentials: config.credentials,
    });

    // Connect and verify
    console.log("[SyncEngine] Calling ewsClient.connect()...");
    var result = await this.ewsClient.connect();
    console.log("[SyncEngine] EWS connected:", result.ewsUrl);

    // Ensure we have a local calendar
    console.log("[SyncEngine] Ensuring local calendar exists...");
    await this._ensureLocalCalendar(config.calendarName || "Exchange Calendar");
    console.log("[SyncEngine] Local calendar ready, id=", this.calendarId);

    // Listen for local changes (two-way sync)
    this._onLocalChange = (changeInfo) => {
      if (this._suppressLocalEvents) return;
      if (changeInfo.calendarId !== this.calendarId) return;
      delete this.conflicts[changeInfo.itemId];
      this._localChangesQueue.push(changeInfo);
      this._saveState().catch((e) => {
        this.lastError = "Could not persist local calendar change: " + e.message;
      });
    };
    browser.CalendarAPI.onItemChanged.addListener(this._onLocalChange);
    this._initialized = true;

    console.log("[SyncEngine] Initialization complete");
    return { success: true, calendarId: this.calendarId };
  }

  /**
   * Ensure a local calendar exists for synced items.
   */
  async _ensureLocalCalendar(name) {
    console.log("[SyncEngine] _ensureLocalCalendar, existing calendarId=", this.calendarId);
    if (this.calendarId) {
      // Verify it still exists
      console.log("[SyncEngine] Checking if existing calendar still exists...");
      var calendars = await browser.CalendarAPI.getCalendars();
      console.log("[SyncEngine] Got", calendars.length, "calendars");
      var exists = calendars.some(function (c) { return c.id === this.calendarId; }.bind(this));
      if (exists) {
        console.log("[SyncEngine] Existing calendar found");
        return;
      }
      console.log("[SyncEngine] Existing calendar not found, will create new");
    }

    // Create new calendar for this account
    console.log("[SyncEngine] Creating new calendar:", name);
    const cal = await browser.CalendarAPI.createCalendar(name, "#0078D4");
    this.calendarId = cal.id;
    await browser.storage.local.set({ calendarId: this.calendarId });
    console.log("[SyncEngine] Created local calendar:", cal.id);
  }

  /**
   * Save sync state to storage.
   */
  async _saveState() {
    await browser.storage.local.set({
      syncState: this.syncState,
      itemMap: this.itemMap,
      reverseMap: this.reverseMap,
      localChangesQueue: this._localChangesQueue,
      conflicts: this.conflicts,
    });
  }

  /**
   * Perform a full initial sync.
   */
  async fullSync() {
    if (this._syncing) {
      console.log("[SyncEngine] Sync already in progress, skipping");
      return { skipped: true };
    }

    this._syncing = true;
    this._suppressLocalEvents = true;

    try {
      console.log("[SyncEngine] Starting full sync...");

      const now = new Date();
      const start = new Date(
        now.getTime() - DEFAULT_SYNC_RANGE_PAST_DAYS * 24 * 60 * 60 * 1000
      );
      const end = new Date(
        now.getTime() + DEFAULT_SYNC_RANGE_FUTURE_DAYS * 24 * 60 * 60 * 1000
      );

      // Capture the cursor before the CalendarView import. This ensures
      // changes that occur during the import remain in the next pull.
      const baselineSyncState = await this.ewsClient.establishSyncState();

      // Fetch all items from Exchange (summary only)
      const ewsSummaries = await this.ewsClient.findCalendarItems(start, end);
      console.log(`[SyncEngine] Found ${ewsSummaries.length} items on Exchange`);

      // Fetch full details (body, attendees, etc.) in batches
      const BATCH_SIZE = 50;
      const fullItemsMap = {};
      for (let i = 0; i < ewsSummaries.length; i += BATCH_SIZE) {
        const batch = ewsSummaries.slice(i, i + BATCH_SIZE);
        const itemIds = batch
          .filter((s) => !s.isCancelled)
          .map((s) => ({ id: s.id, changeKey: s.changeKey }));
        if (itemIds.length > 0) {
          const fullItems = await this.ewsClient.getItems(itemIds);
          for (const fi of fullItems) {
            fullItemsMap[fi.id] = fi;
          }
        }
      }
      console.log(`[SyncEngine] Fetched full details for ${Object.keys(fullItemsMap).length} items`);

      // Merge: use full details where available, fall back to summary
      const ewsItems = ewsSummaries.map((s) => fullItemsMap[s.id] || s);

      // Get current local items
      console.log("[SyncEngine] Getting local items...");
      const localItems = await browser.CalendarAPI.getItems(
        this.calendarId,
        "events"
      );
      console.log("[SyncEngine] Got", localItems.length, "local items");
      const localItemsById = {};
      for (const item of localItems) {
        localItemsById[item.id] = item;
      }

      // Build a UID lookup from existing local items to handle re-sync after
      // itemMap loss (e.g. reinstall). UIDs follow pattern: ews-HASH@exchange-cal-sync
      const localItemsByUid = {};
      for (const item of localItems) {
        if (item.icsData) {
          const uidMatch = item.icsData.match(/^UID:(.+)$/m);
          if (uidMatch) {
            localItemsByUid[uidMatch[1].trim()] = item;
          }
        }
      }

      let created = 0;
      let updated = 0;
      let deleted = 0;

      // Track which EWS items we've seen
      const seenEwsIds = new Set();
      const calendarViewIds = new Set(ewsSummaries.map((item) => item.id));
      let itemIndex = 0;

      for (const ewsItem of ewsItems) {
        itemIndex++;

        // Skip cancelled items - treat as if they don't exist
        if (ewsItem.isCancelled) {
          console.log("[SyncEngine] Skipping cancelled item:", ewsItem.subject);
          continue;
        }

        seenEwsIds.add(ewsItem.id);
        const icsData = ewsItemToIcs(ewsItem);
        const mapping = this.itemMap[ewsItem.id];

        if (!mapping) {
          // Check if local item already exists with matching UID (mapping was lost)
          const expectedUid = ewsItem.id
            ? stableEwsUid(ewsItem.id)
            : null;
          const existingLocal = expectedUid ? localItemsByUid[expectedUid] : null;

          if (existingLocal) {
            // Rebuild mapping from existing local item
            console.log("[SyncEngine] Rebuilding mapping for:", ewsItem.subject);
            this.itemMap[ewsItem.id] = {
              localItemId: existingLocal.id,
              changeKey: ewsItem.changeKey,
              lastModified: ewsItem.lastModified,
            };
            this.reverseMap[existingLocal.id] = ewsItem.id;
            // Update it with latest data
            try {
              await browser.CalendarAPI.updateItem(
                this.calendarId,
                existingLocal.id,
                icsData
              );
              updated++;
            } catch (e) {
              // Non-critical - mapping is rebuilt either way
            }
          } else {
            // Genuinely new item -> create locally
            try {
              console.log("[SyncEngine] Adding item", itemIndex, "/", ewsItems.length, ":", ewsItem.subject);
              const result = await browser.CalendarAPI.addItem(
                this.calendarId,
                icsData
              );
              console.log("[SyncEngine] Item added, localId=", result.id);
              this.itemMap[ewsItem.id] = {
                localItemId: result.id,
                changeKey: ewsItem.changeKey,
                lastModified: ewsItem.lastModified,
              };
              this.reverseMap[result.id] = ewsItem.id;
              created++;
            } catch (e) {
              console.error(
                `[SyncEngine] Failed to create local item for "${ewsItem.subject}":`,
                e
              );
            }
          }
        } else if (mapping.changeKey !== ewsItem.changeKey) {
          // Item changed on Exchange -> update locally
          try {
            await browser.CalendarAPI.updateItem(
              this.calendarId,
              mapping.localItemId,
              icsData
            );
            mapping.changeKey = ewsItem.changeKey;
            mapping.lastModified = ewsItem.lastModified;
            updated++;
          } catch (e) {
            console.error(
              `[SyncEngine] Failed to update local item for "${ewsItem.subject}":`,
              e
            );
          }
        }
      }

      // Delete local items that no longer exist on Exchange
      for (const [ewsId, mapping] of Object.entries(this.itemMap)) {
        // CalendarView only covers the configured range. Never delete an
        // event merely because it is outside that range.
        if (calendarViewIds.has(ewsId) && !seenEwsIds.has(ewsId)) {
          try {
            await browser.CalendarAPI.deleteItem(
              this.calendarId,
              mapping.localItemId
            );
            delete this.reverseMap[mapping.localItemId];
            delete this.itemMap[ewsId];
            deleted++;
          } catch (e) {
            // Item may already be deleted locally
            delete this.reverseMap[mapping.localItemId];
            delete this.itemMap[ewsId];
          }
        }
      }

      // Seed SyncFolderItems after the range import. Without this cursor the
      // next periodic refresh would repeat the complete CalendarView import.
      this.syncState = baselineSyncState;
      await this._saveState();

      const result = { created, updated, deleted, total: ewsItems.length };
      console.log("[SyncEngine] Full sync complete:", result);
      return result;
    } finally {
      this._syncing = false;
      this._suppressLocalEvents = false;
    }
  }

  /**
   * Perform an incremental sync using SyncFolderItems.
   */
  async incrementalSync() {
    if (this._syncing) {
      console.log("[SyncEngine] Sync already in progress, skipping");
      return { skipped: true };
    }

    this._syncing = true;

    try {
      // First push local changes to Exchange
      await this._pushLocalChanges();

      // Save state after push so we don't re-push if pull fails
      await this._saveState();

      // Then pull remote changes
      this._suppressLocalEvents = true;
      const result = await this._pullRemoteChanges();
      this._suppressLocalEvents = false;

      await this._saveState();
      return result;
    } finally {
      this._syncing = false;
      this._suppressLocalEvents = false;
    }
  }

  /**
   * Push queued local changes to Exchange.
   */
  async _pushLocalChanges() {
    const queue = [...this._localChangesQueue];
    this._localChangesQueue = [];

    if (queue.length === 0) return;
    console.log(`[SyncEngine] Pushing ${queue.length} local changes to Exchange...`);

    let pushed = 0;
    let errors = 0;
    let firstError = null;

    for (const change of queue) {
      try {
        const ewsId = this.reverseMap[change.itemId];

        if (change.changeType === "added" && !ewsId) {
          // New local item -> create on Exchange
          if (change.icsData) {
            const fields = icsToEwsFields(change.icsData);
            const result = await this.ewsClient.createItem(fields);
            this.itemMap[result.id] = {
              localItemId: change.itemId,
              changeKey: result.changeKey,
              lastModified: new Date().toISOString(),
            };
            this.reverseMap[change.itemId] = result.id;
            pushed++;
          }
        } else if (change.changeType === "modified" && ewsId) {
          // Modified local item -> update on Exchange
          const mapping = this.itemMap[ewsId];
          if (mapping && change.icsData) {
            const fields = icsToEwsFields(change.icsData);
            const result = await this.ewsClient.updateItem(
              ewsId,
              mapping.changeKey,
              fields
            );
            mapping.changeKey = result.changeKey;
            mapping.lastModified = new Date().toISOString();
            pushed++;
          }
        } else if (change.changeType === "deleted" && ewsId) {
          // Deleted local item -> delete on Exchange
          const mapping = this.itemMap[ewsId];
          if (mapping) {
            await this.ewsClient.deleteItem(ewsId, mapping.changeKey);
            delete this.itemMap[ewsId];
            delete this.reverseMap[change.itemId];
            pushed++;
          }
        }
      } catch (e) {
        console.error("[SyncEngine] Failed to push local change:", e);
        errors++;
        if (!firstError) firstError = e;
        const ewsId = this.reverseMap[change.itemId];
        if (ewsId && isEwsConflictError(e)) {
          await this._recordConflict(change, ewsId, e);
        } else {
          // Re-queue transient failures for retry.
          this._localChangesQueue.push(change);
        }
      }
    }

    if (pushed > 0 || errors > 0) {
      console.log(
        `[SyncEngine] Pushed ${pushed} local changes, ${errors} errors`
      );
    }
    if (firstError) {
      await this._saveState();
      throw firstError;
    }
  }

  async _recordConflict(change, ewsId, error) {
    let remote = null;
    try {
      const items = await this.ewsClient.getItems([{ id: ewsId }]);
      remote = items[0] || null;
    } catch (refreshError) {
      console.error("[SyncEngine] Could not refresh conflicting item:", refreshError);
    }
    this.conflicts[change.itemId] = {
      ewsId,
      error: error.message,
      remoteChangeKey: remote && remote.changeKey,
      remoteLastModified: remote && remote.lastModified,
      detectedAt: new Date().toISOString(),
    };
  }

  /**
   * Pull remote changes from Exchange using SyncFolderItems.
   */
  async _pullRemoteChanges() {
    // If no syncState, do a full sync instead
    if (!this.syncState) {
      this._syncing = false;
      return await this.fullSync();
    }

    let created = 0;
    let updated = 0;
    let deleted = 0;
    let hasMore = true;

    while (hasMore) {
      const cursor = this.syncState;
      const changes = await this.ewsClient.syncFolderItems(cursor);

      // Handle created items
      if (changes.created.length > 0) {
        // Filter out items we already have mapped (e.g. from our own push)
        const newCreates = changes.created.filter((c) => !this.itemMap[c.id]);
        const ownCreates = changes.created.filter((c) => this.itemMap[c.id]);

        // Update changeKeys for items we pushed (server may assign new ones)
        for (const c of ownCreates) {
          const mapping = this.itemMap[c.id];
          if (mapping) {
            mapping.changeKey = c.changeKey;
          }
        }

        if (newCreates.length > 0) {
          const itemIds = newCreates.map((c) => ({
            id: c.id,
            changeKey: c.changeKey,
          }));
          const fullItems = await this.ewsClient.getItems(itemIds);

          for (const ewsItem of fullItems) {
            // If cancelled, skip creating locally
            if (ewsItem.isCancelled) {
              console.log("[SyncEngine] Skipping cancelled new item:", ewsItem.subject);
              continue;
            }
            const icsData = ewsItemToIcs(ewsItem);
            try {
              const result = await browser.CalendarAPI.addItem(
                this.calendarId,
                icsData
              );
              this.itemMap[ewsItem.id] = {
                localItemId: result.id,
                changeKey: ewsItem.changeKey,
                lastModified: ewsItem.lastModified,
              };
              this.reverseMap[result.id] = ewsItem.id;
              created++;
            } catch (e) {
              console.error("[SyncEngine] Failed to create local item:", e);
              throw e;
            }
          }
        }
      }

      // Handle updated items
      if (changes.updated.length > 0) {
        // Filter out updates where we already have the latest changeKey (our own push)
        const realUpdates = changes.updated.filter((c) => {
          const mapping = this.itemMap[c.id];
          return !mapping || mapping.changeKey !== c.changeKey;
        });

        if (realUpdates.length > 0) {
          const itemIds = realUpdates.map((c) => ({
            id: c.id,
            changeKey: c.changeKey,
          }));
          const fullItems = await this.ewsClient.getItems(itemIds);

        for (const ewsItem of fullItems) {
          const mapping = this.itemMap[ewsItem.id];

          // If item is now cancelled, delete it locally
          if (ewsItem.isCancelled) {
            console.log("[SyncEngine] Item cancelled on Exchange, deleting locally:", ewsItem.subject);
            if (mapping) {
              if (this.conflicts[mapping.localItemId]) {
                this.conflicts[mapping.localItemId].remoteDeleted = true;
                continue;
              }
              try {
                await browser.CalendarAPI.deleteItem(
                  this.calendarId,
                  mapping.localItemId
                );
              } catch (e) {
                throw e;
              }
              delete this.reverseMap[mapping.localItemId];
              delete this.itemMap[ewsItem.id];
              deleted++;
            }
            continue;
          }

          if (mapping) {
            if (this.conflicts[mapping.localItemId]) {
              this.conflicts[mapping.localItemId].remoteChangeKey = ewsItem.changeKey;
              this.conflicts[mapping.localItemId].remoteLastModified = ewsItem.lastModified;
              continue;
            }
            const icsData = ewsItemToIcs(ewsItem);
            try {
              await browser.CalendarAPI.updateItem(
                this.calendarId,
                mapping.localItemId,
                icsData
              );
              mapping.changeKey = ewsItem.changeKey;
              mapping.lastModified = ewsItem.lastModified;
              updated++;
            } catch (e) {
              console.error("[SyncEngine] Failed to update local item:", e);
              throw e;
            }
          } else {
            // Item not mapped yet - treat as creation
            const icsData = ewsItemToIcs(ewsItem);
            try {
              const result = await browser.CalendarAPI.addItem(
                this.calendarId,
                icsData
              );
              this.itemMap[ewsItem.id] = {
                localItemId: result.id,
                changeKey: ewsItem.changeKey,
                lastModified: ewsItem.lastModified,
              };
              this.reverseMap[result.id] = ewsItem.id;
              created++;
            } catch (e) {
              console.error("[SyncEngine] Failed to create local item:", e);
              throw e;
            }
          }
        }
        }
      }

      // Handle deleted items
      for (const del of changes.deleted) {
        const mapping = this.itemMap[del.id];
        if (mapping) {
          if (this.conflicts[mapping.localItemId]) {
            this.conflicts[mapping.localItemId].remoteDeleted = true;
            continue;
          }
          try {
            await browser.CalendarAPI.deleteItem(
              this.calendarId,
              mapping.localItemId
            );
          } catch (e) {
            throw e;
          }
          delete this.reverseMap[mapping.localItemId];
          delete this.itemMap[del.id];
          deleted++;
        }
      }

      // Advance the cursor only after every change in this page has been
      // applied. A failed page is retried on the next refresh.
      this.syncState = changes.syncState;
      hasMore = !changes.includesLastItem;
    }

    const result = { created, updated, deleted, incremental: true };
    console.log("[SyncEngine] Incremental sync complete:", result);
    return result;
  }

  /**
   * Reset sync state (forces full re-sync on next sync).
   */
  async resetSync() {
    this.syncState = null;
    this.itemMap = {};
    this.reverseMap = {};
    this._localChangesQueue = [];
    this.conflicts = {};
    await this._saveState();
    console.log("[SyncEngine] Sync state reset");
  }

  /**
   * Get current sync status.
   */
  getStatus() {
    return {
      connected: this.ewsClient?._verified || false,
      calendarId: this.calendarId,
      hasSyncState: !!this.syncState,
      mappedItems: Object.keys(this.itemMap).length,
      pendingLocalChanges: this._localChangesQueue.length,
      conflicts: Object.keys(this.conflicts),
      lastError: this.lastError,
      syncing: this._syncing,
    };
  }
}

// Globals: SyncEngine, ewsItemToIcs, icsToEwsFields, DEFAULT_SYNC_INTERVAL_MINUTES
