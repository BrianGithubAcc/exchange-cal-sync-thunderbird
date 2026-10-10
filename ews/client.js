console.log("[ExCalSync] client.js loading");
/**
 * EWS (Exchange Web Services) Client.
 * Handles authenticated SOAP communication with Exchange servers.
 * Uses OAuth 2.0 Bearer tokens for Exchange Online.
 */

// Depends on: ews/xml-builder.js

const DEFAULT_EWS_URL = "https://outlook.office365.com/EWS/Exchange.asmx";

function validateEwsUrl(url) {
  const value = url || DEFAULT_EWS_URL;
  let parsed;
  try {
    parsed = new URL(value);
  } catch (e) {
    throw new Error("Invalid EWS URL");
  }
  if (
    parsed.protocol !== "https:" ||
    parsed.hostname !== "outlook.office365.com" ||
    parsed.port || parsed.username || parsed.password ||
    parsed.search || parsed.hash ||
    parsed.pathname !== "/EWS/Exchange.asmx"
  ) {
    throw new Error("EWS URL must be https://outlook.office365.com/EWS/Exchange.asmx");
  }
  return DEFAULT_EWS_URL;
}

/**
 * Parse XML response text into a DOM Document.
 */
function parseXml(xmlText) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, "text/xml");
  const parseError = doc.querySelector("parsererror");
  if (parseError) {
    throw new Error(`XML parse error: ${parseError.textContent}`);
  }
  return doc;
}

/**
 * Extract text content from an element by local tag name.
 */
function getTagText(parent, tagName) {
  // Use local-name() matching since namespaces vary
  const el = parent.querySelector(tagName);
  if (el) return el.textContent.trim();
  // Fallback: search all elements for local name match
  const all = parent.getElementsByTagName("*");
  for (const node of all) {
    if (node.localName === tagName) {
      return node.textContent.trim();
    }
  }
  return null;
}

/**
 * Get all elements matching a local tag name.
 */
function getElementsByLocalName(parent, localName) {
  const results = [];
  const all = parent.getElementsByTagName("*");
  for (const node of all) {
    if (node.localName === localName) {
      results.push(node);
    }
  }
  return results;
}

/**
 * Check EWS SOAP response for errors.
 */
function checkSoapFault(doc) {
  const faults = getElementsByLocalName(doc, "Fault");
  if (faults.length > 0) {
    const faultString =
      getTagText(faults[0], "faultstring") || "Unknown SOAP Fault";
    const faultCode =
      getTagText(faults[0], "faultcode") || "Unknown";
    throw new Error(`EWS SOAP Fault [${faultCode}]: ${faultString}`);
  }

  // Check ResponseCode for EWS-level errors
  const responseMessages = getElementsByLocalName(doc, "ResponseCode");
  for (const rm of responseMessages) {
    const code = rm.textContent.trim();
    if (code !== "NoError") {
      const msgText =
        getTagText(rm.parentElement, "MessageText") || code;
      throw new Error(`EWS Error [${code}]: ${msgText}`);
    }
  }
}

/**
 * EWS Client class.
 */
class EwsClient {
  /**
   * @param {object} config
   * @param {string} config.email - User email address
   * @param {string} [config.ewsUrl] - Microsoft 365 EWS endpoint URL
   * @param {object} config.credentials - Auth credentials
   * @param {string} [config.credentials.token] - For OAuth bearer token
   */
  constructor(config) {
    this.email = config.email;
    this.ewsUrl = validateEwsUrl(config.ewsUrl || DEFAULT_EWS_URL);
    this.credentials = config.credentials;
    this._verified = false;
  }

  /**
   * Build Authorization header based on credential type.
   */
  _getAuthHeaders() {
    const headers = {
      "Content-Type": "text/xml; charset=utf-8",
    };

    if (this.credentials.type === "oauth") {
      if (!this.credentials.token) {
        throw new Error("OAuth requires an access token");
      }
      headers["Authorization"] = `Bearer ${this.credentials.token}`;
    } else {
      throw new Error(`Unsupported auth type: ${this.credentials.type}`);
    }

    return headers;
  }

  /**
   * Send a SOAP request to the EWS endpoint.
   */
  async _sendRequest(soapXml) {
    if (!this.ewsUrl) {
      throw new Error("EWS URL not configured. Call connect() first.");
    }

    const headers = this._getAuthHeaders();

    let response = await this._fetch(this.ewsUrl, headers, soapXml);

    if (response.status === 401 && this.credentials.type === "oauth" && this._refreshTokenFn) {
      // Try to refresh the OAuth token and retry once
      console.log("[EWS] Got 401, attempting token refresh...");
      const newToken = await this._refreshTokenFn();
      if (newToken) {
        this.credentials.token = newToken;
        const retryHeaders = this._getAuthHeaders();
        response = await this._fetch(this.ewsUrl, retryHeaders, soapXml);
      }
    }

    if (response.status === 401) {
      throw new Error(
          "Authentication failed (HTTP 401). Sign in again with Microsoft OAuth."
      );
    }

    if (response.status === 403) {
      throw new Error(
        "Access forbidden (HTTP 403). Your account may not have EWS access enabled."
      );
    }

    if (!response.ok) {
      throw new Error(
        `EWS request failed: HTTP ${response.status} ${response.statusText}`
      );
    }

    const responseText = await response.text();
    const doc = parseXml(responseText);
    checkSoapFault(doc);
    return doc;
  }

  async _fetch(url, headers, body) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000);
    try {
      return await fetch(url, {
        method: "POST",
        headers,
        body,
        redirect: "error",
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timeout);
    }
  }

  /**
   * Connect to Exchange: discover EWS URL if needed and verify connectivity.
   */
  async connect() {
    this.ewsUrl = validateEwsUrl(this.ewsUrl);

    // Verify the default calendar folder before reading events.
    console.log("[EWS] Verifying connection...");
    try {
      await this.getCalendarFolder();
      this._verified = true;
      console.log("[EWS] Connection verified successfully");
      return { success: true, ewsUrl: this.ewsUrl };
    } catch (e) {
      this._verified = false;
      throw new Error(`Connection verification failed: ${e.message}`);
    }
  }

  async getCalendarFolder() {
    const doc = await this._sendRequest(buildGetCalendarFolderRequest());
    const folder = getElementsByLocalName(doc, "CalendarFolder")[0] ||
      getElementsByLocalName(doc, "Folder")[0];
    if (!folder) throw new Error("EWS did not return the default calendar folder");
    const id = getElementsByLocalName(folder, "FolderId")[0];
    return {
      id: id && id.getAttribute("Id"),
      changeKey: id && id.getAttribute("ChangeKey"),
      displayName: getTagText(folder, "DisplayName") || "Calendar",
    };
  }

  /**
   * Find calendar items in a date range.
   * Returns an array of calendar item summaries.
   */
  async findCalendarItems(startDate, endDate) {
    const requestXml = buildFindItemRequest(startDate, endDate);
    const doc = await this._sendRequest(requestXml);

    const items = [];
    const calendarItems = getElementsByLocalName(doc, "CalendarItem");

    for (const item of calendarItems) {
      const itemIdEl = getElementsByLocalName(item, "ItemId")[0];
      if (!itemIdEl) continue;

      items.push({
        id: itemIdEl.getAttribute("Id"),
        changeKey: itemIdEl.getAttribute("ChangeKey"),
        subject: getTagText(item, "Subject") || "(No Subject)",
        start: getTagText(item, "Start"),
        end: getTagText(item, "End"),
        location: getTagText(item, "Location"),
        isAllDay: getTagText(item, "IsAllDayEvent") === "true",
        body: getTagText(item, "Body"),
        sensitivity: getTagText(item, "Sensitivity") || "Normal",
        importance: getTagText(item, "Importance") || "Normal",
        lastModified: getTagText(item, "LastModifiedTime"),
        isCancelled: getTagText(item, "IsCancelled") === "true",
        hasAttachments: getTagText(item, "HasAttachments") === "true",
        categories: (() => {
          const cats = [];
          const catEls = getElementsByLocalName(item, "String");
          const catParent = getElementsByLocalName(item, "Categories")[0];
          if (catParent) {
            for (const c of getElementsByLocalName(catParent, "String")) {
              cats.push(c.textContent.trim());
            }
          }
          return cats;
        })(),
        reminderMinutes: parseInt(
          getTagText(item, "ReminderMinutesBeforeStart") || "0",
          10
        ),
        reminderSet: getTagText(item, "ReminderIsSet") === "true",
      });
    }

    return items;
  }

  /**
   * Get full details of specific items by their IDs.
   */
  async getItems(itemIds) {
    if (itemIds.length === 0) return [];

    const requestXml = buildGetItemRequest(itemIds);
    const doc = await this._sendRequest(requestXml);

    const items = [];
    const calendarItems = getElementsByLocalName(doc, "CalendarItem");

    for (const item of calendarItems) {
      const itemIdEl = getElementsByLocalName(item, "ItemId")[0];
      if (!itemIdEl) continue;

      // Parse attendees
      const attendees = [];
      const reqAttendees = getElementsByLocalName(item, "RequiredAttendees")[0];
      if (reqAttendees) {
        for (const a of getElementsByLocalName(reqAttendees, "Attendee")) {
          const email = getTagText(a, "EmailAddress");
          const name = getTagText(a, "Name");
          if (email) attendees.push({ email, name, type: "required" });
        }
      }
      const optAttendees = getElementsByLocalName(item, "OptionalAttendees")[0];
      if (optAttendees) {
        for (const a of getElementsByLocalName(optAttendees, "Attendee")) {
          const email = getTagText(a, "EmailAddress");
          const name = getTagText(a, "Name");
          if (email) attendees.push({ email, name, type: "optional" });
        }
      }

      // Parse organizer
      const organizerEl = getElementsByLocalName(item, "Organizer")[0];
      let organizer = null;
      if (organizerEl) {
        const orgMailbox = getElementsByLocalName(organizerEl, "Mailbox")[0];
        if (orgMailbox) {
          organizer = {
            name: getTagText(orgMailbox, "Name") || "",
            email: getTagText(orgMailbox, "EmailAddress") || "",
          };
        }
      }

      // Capture body and its type (HTML or Text)
      const bodyEl = getElementsByLocalName(item, "Body")[0];
      const bodyType = bodyEl ? (bodyEl.getAttribute("BodyType") || "Text") : "Text";
      const bodyContent = bodyEl ? bodyEl.textContent.trim() : "";

      items.push({
        id: itemIdEl.getAttribute("Id"),
        changeKey: itemIdEl.getAttribute("ChangeKey"),
        subject: getTagText(item, "Subject") || "(No Subject)",
        start: getTagText(item, "Start"),
        end: getTagText(item, "End"),
        location: getTagText(item, "Location"),
        isAllDay: getTagText(item, "IsAllDayEvent") === "true",
        body: bodyContent,
        bodyType: bodyType,
        sensitivity: getTagText(item, "Sensitivity") || "Normal",
        importance: getTagText(item, "Importance") || "Normal",
        lastModified: getTagText(item, "LastModifiedTime"),
        isCancelled: getTagText(item, "IsCancelled") === "true",
        attendees,
        organizer,
      });
    }

    return items;
  }

  /**
   * Sync folder items incrementally.
   * Returns changes since the last sync state.
   */
  async syncFolderItems(syncState = null) {
    const requestXml = buildSyncFolderItemsRequest(syncState);
    const doc = await this._sendRequest(requestXml);

    const newSyncState = getTagText(doc.documentElement, "SyncState");
    const includesLastItem =
      getTagText(doc.documentElement, "IncludesLastItemInRange") === "true";

    const changes = {
      created: [],
      updated: [],
      deleted: [],
      syncState: newSyncState,
      includesLastItem,
    };

    // Parse Create changes
    for (const el of getElementsByLocalName(doc, "Create")) {
      const itemId = getElementsByLocalName(el, "ItemId")[0];
      if (itemId) {
        changes.created.push({
          id: itemId.getAttribute("Id"),
          changeKey: itemId.getAttribute("ChangeKey"),
          subject: getTagText(el, "Subject"),
        });
      }
    }

    // Parse Update changes
    for (const el of getElementsByLocalName(doc, "Update")) {
      const itemId = getElementsByLocalName(el, "ItemId")[0];
      if (itemId) {
        changes.updated.push({
          id: itemId.getAttribute("Id"),
          changeKey: itemId.getAttribute("ChangeKey"),
          subject: getTagText(el, "Subject"),
        });
      }
    }

    // Parse Delete changes
    for (const el of getElementsByLocalName(doc, "Delete")) {
      const itemId = getElementsByLocalName(el, "ItemId")[0];
      if (itemId) {
        changes.deleted.push({
          id: itemId.getAttribute("Id"),
          changeKey: itemId.getAttribute("ChangeKey"),
        });
      }
    }

    return changes;
  }

  /**
   * Establish an incremental-sync cursor without applying the returned
   * changes. This is used after the initial CalendarView import so later
   * refreshes do not repeatedly perform a full sync.
   */
  async establishSyncState() {
    let syncState = null;
    let complete = false;
    while (!complete) {
      const changes = await this.syncFolderItems(syncState);
      syncState = changes.syncState;
      complete = changes.includesLastItem;
      if (!syncState) {
        throw new Error("EWS did not return a SyncState for the calendar");
      }
    }
    return syncState;
  }

  /**
   * Create a new calendar event on Exchange.
   */
  async createItem(event) {
    const requestXml = buildCreateItemRequest(event);
    const doc = await this._sendRequest(requestXml);

    const itemIdEl = getElementsByLocalName(doc, "ItemId")[0];
    if (!itemIdEl) {
      throw new Error("CreateItem succeeded but no ItemId returned");
    }

    return {
      id: itemIdEl.getAttribute("Id"),
      changeKey: itemIdEl.getAttribute("ChangeKey"),
    };
  }

  /**
   * Update an existing calendar event on Exchange.
   */
  async updateItem(itemId, changeKey, updates) {
    const requestXml = buildUpdateItemRequest(itemId, changeKey, updates);
    const doc = await this._sendRequest(requestXml);

    const itemIdEl = getElementsByLocalName(doc, "ItemId")[0];
    return {
      id: itemIdEl ? itemIdEl.getAttribute("Id") : itemId,
      changeKey: itemIdEl
        ? itemIdEl.getAttribute("ChangeKey")
        : changeKey,
    };
  }

  /**
   * Delete a calendar event on Exchange.
   */
  async deleteItem(itemId, changeKey) {
    const requestXml = buildDeleteItemRequest(itemId, changeKey);
    await this._sendRequest(requestXml);
    return { success: true };
  }

  /**
   * Respond to a meeting invitation (Accept/Tentative/Decline).
   * @param {string} itemId - CalendarItem ItemId
   * @param {string} changeKey - CalendarItem ChangeKey
   * @param {string} responseType - "accept", "tentative", or "decline"
   */
  async respondToMeeting(itemId, changeKey, responseType) {
    const requestXml = buildMeetingResponseRequest(itemId, changeKey, responseType);
    await this._sendRequest(requestXml);
    return { success: true };
  }

  /**
   * Find a calendar item by subject within a date range.
   * Returns matching items with their ItemId and ChangeKey.
   */
  async findItemBySubject(subject, startDate, endDate) {
    const requestXml = buildFindItemBySubjectRequest(subject, startDate, endDate);
    const doc = await this._sendRequest(requestXml);

    const items = [];
    const calendarItems = getElementsByLocalName(doc, "CalendarItem");
    for (const item of calendarItems) {
      const itemIdEl = getElementsByLocalName(item, "ItemId")[0];
      if (!itemIdEl) continue;
      items.push({
        id: itemIdEl.getAttribute("Id"),
        changeKey: itemIdEl.getAttribute("ChangeKey"),
        subject: getTagText(item, "Subject") || "",
        start: getTagText(item, "Start"),
        end: getTagText(item, "End"),
        location: getTagText(item, "Location"),
        organizer: getTagText(item, "Name"), // Organizer > Mailbox > Name
      });
    }
    return items;
  }
}

// Globals: EwsClient, parseXml, getTagText, getElementsByLocalName
