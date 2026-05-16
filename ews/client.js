console.log("[ExCalSync] client.js loading");
/**
 * EWS (Exchange Web Services) Client.
 * Handles authenticated SOAP communication with Exchange servers.
 * Supports Basic Auth and OAuth 2.0 Bearer tokens.
 */

// Depends on: ews/xml-builder.js, ews/autodiscover.js

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
   * @param {string} [config.ewsUrl] - EWS endpoint URL (auto-discovered if not provided)
   * @param {object} config.credentials - Auth credentials
   * @param {string} config.credentials.type - 'basic' or 'oauth'
   * @param {string} [config.credentials.username] - For basic auth
   * @param {string} [config.credentials.password] - For basic auth
   * @param {string} [config.credentials.token] - For OAuth bearer token
   */
  constructor(config) {
    this.email = config.email;
    this.ewsUrl = config.ewsUrl || null;
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

    if (this.credentials.type === "basic") {
      if (!this.credentials.username || !this.credentials.password) {
        throw new Error("Basic auth requires username and password");
      }
      const encoded = btoa(
        String.fromCharCode(
          ...new TextEncoder().encode(
            `${this.credentials.username}:${this.credentials.password}`
          )
        )
      );
      headers["Authorization"] = `Basic ${encoded}`;
    } else if (this.credentials.type === "oauth") {
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

    let response = await fetch(this.ewsUrl, {
      method: "POST",
      headers,
      body: soapXml,
    });

    if (response.status === 401 && this.credentials.type === "oauth" && this._refreshTokenFn) {
      // Try to refresh the OAuth token and retry once
      console.log("[EWS] Got 401, attempting token refresh...");
      const newToken = await this._refreshTokenFn();
      if (newToken) {
        this.credentials.token = newToken;
        const retryHeaders = this._getAuthHeaders();
        response = await fetch(this.ewsUrl, {
          method: "POST",
          headers: retryHeaders,
          body: soapXml,
        });
      }
    }

    if (response.status === 401) {
      throw new Error(
        "Authentication failed (HTTP 401). Check your credentials. " +
          "If using Basic Auth with Microsoft 365, it may be disabled for your tenant. " +
          "Consider using OAuth 2.0 instead."
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

  /**
   * Connect to Exchange: discover EWS URL if needed and verify connectivity.
   */
  async connect() {
    // Step 1: Discover EWS URL if not provided
    if (!this.ewsUrl) {
      console.log("[EWS] Starting autodiscover for:", this.email);
      try {
        this.ewsUrl = await autodiscover(this.email, this.credentials);
      } catch (e) {
        console.error("[EWS] Autodiscover failed:", e.message);
        // For M365, try the standard URL as fallback
        console.log("[EWS] Trying standard M365 EWS URL as fallback...");
        this.ewsUrl = "https://outlook.office365.com/EWS/Exchange.asmx";
      }
      console.log("[EWS] Using URL:", this.ewsUrl);
    } else {
      this.ewsUrl = validateEwsUrl(this.ewsUrl);
      console.log("[EWS] Using provided URL:", this.ewsUrl);
    }

    // Step 2: Verify connectivity by fetching a small date range
    console.log("[EWS] Verifying connection...");
    const now = new Date();
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    try {
      await this.findCalendarItems(weekAgo, now);
      this._verified = true;
      console.log("[EWS] Connection verified successfully");
      return { success: true, ewsUrl: this.ewsUrl };
    } catch (e) {
      this._verified = false;
      throw new Error(`Connection verification failed: ${e.message}`);
    }
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
