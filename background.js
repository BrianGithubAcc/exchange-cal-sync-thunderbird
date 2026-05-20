console.log("[ExCalSync] background.js loading");
/**
 * Background script for Exchange Calendar Sync.
 * Supports multiple Exchange accounts, each with its own sync engine.
 */

// Depends on: sync/engine.js (SyncEngine, DEFAULT_SYNC_INTERVAL_MINUTES)
// Depends on: sync/imip.js (parseICalendar, isCalendarInvite, imipGetDomain)
// Depends on: ews/client.js (EwsClient)

var ALARM_PREFIX = "excalsync-";
var MS_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c";
var EWS_SCOPE = "https://outlook.office365.com/EWS.AccessAsUser.All offline_access";

// Map of accountId -> { engine, lastResult, lastError }
var accounts = {};

function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
}

// Calendar property keys for persisting account config across reinstalls
var CAL_PROP_PREFIX = "exchange-cal-sync.";
var CAL_PROP_ACCOUNT_ID = CAL_PROP_PREFIX + "accountId";
var CAL_PROP_EMAIL = CAL_PROP_PREFIX + "email";
var CAL_PROP_AUTH_TYPE = CAL_PROP_PREFIX + "authType";
var CAL_PROP_EWS_URL = CAL_PROP_PREFIX + "ewsUrl";
var CAL_PROP_CAL_NAME = CAL_PROP_PREFIX + "calendarName";
var CAL_PROP_SYNC_INTERVAL = CAL_PROP_PREFIX + "syncInterval";

/**
 * Save account metadata to calendar properties (survives addon reinstall).
 */
async function saveAccountToCalendar(calendarId, account) {
  try {
    await browser.CalendarAPI.setCalendarProperty(calendarId, CAL_PROP_ACCOUNT_ID, account.id);
    await browser.CalendarAPI.setCalendarProperty(calendarId, CAL_PROP_EMAIL, account.email);
    await browser.CalendarAPI.setCalendarProperty(calendarId, CAL_PROP_AUTH_TYPE, account.authType || "oauth");
    await browser.CalendarAPI.setCalendarProperty(calendarId, CAL_PROP_EWS_URL, account.ewsUrl || "");
    await browser.CalendarAPI.setCalendarProperty(calendarId, CAL_PROP_CAL_NAME, account.calendarName || "");
    await browser.CalendarAPI.setCalendarProperty(calendarId, CAL_PROP_SYNC_INTERVAL, String(account.syncIntervalMinutes || DEFAULT_SYNC_INTERVAL_MINUTES));
    console.log("[ExCalSync] Saved account metadata to calendar properties:", calendarId);
  } catch (e) {
    console.error("[ExCalSync] Failed to save account to calendar properties:", e);
  }
}

/**
 * Recover accounts from calendar properties when storage is empty (e.g. after reinstall).
 * Returns an array of recovered account objects.
 */
async function recoverAccountsFromCalendars() {
  var recovered = [];
  try {
    var calendars = await browser.CalendarAPI.getCalendars();
    for (var cal of calendars) {
      if (!cal.managed) continue;
      var accountId = await browser.CalendarAPI.getCalendarProperty(cal.id, CAL_PROP_ACCOUNT_ID);
      var email = await browser.CalendarAPI.getCalendarProperty(cal.id, CAL_PROP_EMAIL);
      if (!accountId || !email) continue;

      var authType = await browser.CalendarAPI.getCalendarProperty(cal.id, CAL_PROP_AUTH_TYPE) || "oauth";
      var ewsUrl = await browser.CalendarAPI.getCalendarProperty(cal.id, CAL_PROP_EWS_URL) || "";
      var calName = await browser.CalendarAPI.getCalendarProperty(cal.id, CAL_PROP_CAL_NAME) || cal.name || "Exchange Calendar";
      var syncInterval = parseInt(await browser.CalendarAPI.getCalendarProperty(cal.id, CAL_PROP_SYNC_INTERVAL) || "5", 10);

      var acct = {
        id: accountId,
        email: email,
        authType: authType,
        calendarName: calName,
        ewsUrl: ewsUrl,
        syncIntervalMinutes: syncInterval,
      };

      // Pre-seed the calendarId in storage so the engine finds it
      var calStore = {};
      calStore[acctKey(accountId, "calendarId")] = cal.id;
      await browser.storage.local.set(calStore);

      recovered.push(acct);
      console.log("[ExCalSync] Recovered account from calendar:", email, "id:", accountId, "calendarId:", cal.id);
    }
  } catch (e) {
    console.error("[ExCalSync] Recovery from calendar properties failed:", e);
  }
  return recovered;
}

/**
 * Load all accounts from storage.
 */
async function loadAccounts() {
  var data = await browser.storage.local.get(["accounts"]);
  return data.accounts || [];
}

/**
 * Save accounts list to storage.
 */
async function saveAccounts(accountList) {
  await browser.storage.local.set({ accounts: accountList });
}

/**
 * Get account-scoped storage keys.
 */
function acctKey(accountId, key) {
  return "acct_" + accountId + "_" + key;
}

/**
 * Refresh an OAuth access token using the stored refresh token.
 * Returns the new access token, or null if refresh failed.
 */
async function refreshOAuthToken(accountId) {
  var refreshKey = acctKey(accountId, "oauthRefreshToken");
  var data = await browser.storage.local.get([refreshKey]);
  var refreshToken = data[refreshKey];

  if (!refreshToken) {
    console.log("[ExCalSync] No refresh token for", accountId);
    return null;
  }

  console.log("[ExCalSync] Refreshing OAuth token for", accountId);
  try {
    var response = await fetch(
      "https://login.microsoftonline.com/organizations/oauth2/v2.0/token",
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: "client_id=" + encodeURIComponent(MS_CLIENT_ID) +
              "&grant_type=refresh_token" +
              "&refresh_token=" + encodeURIComponent(refreshToken) +
              "&scope=" + encodeURIComponent(EWS_SCOPE),
      }
    );

    if (!response.ok) {
      var errBody = await response.json().catch(function () { return {}; });
      console.error("[ExCalSync] Token refresh failed:", response.status, errBody.error_description || errBody.error || "");
      return null;
    }

    var tokenData = await response.json();
    var store = {};
    store[acctKey(accountId, "oauthToken")] = tokenData.access_token;
    if (tokenData.refresh_token) {
      store[acctKey(accountId, "oauthRefreshToken")] = tokenData.refresh_token;
    }
    await browser.storage.local.set(store);

    console.log("[ExCalSync] OAuth token refreshed for", accountId);
    return tokenData.access_token;
  } catch (e) {
    console.error("[ExCalSync] Token refresh error for", accountId, ":", e.message);
    return null;
  }
}

/**
 * Start sync for a single account.
 */
async function startAccountSync(account) {
  var id = account.id;
  console.log("[ExCalSync] startAccountSync:", id, account.email);

  try {
    var config = {
      email: account.email,
      ewsUrl: account.ewsUrl || null,
      calendarName: account.calendarName || "Exchange (" + account.email + ")",
      syncIntervalMinutes: account.syncIntervalMinutes || DEFAULT_SYNC_INTERVAL_MINUTES,
      credentials: null,
    };

    if (account.authType === "oauth") {
      // Always try to refresh the token first to avoid 401s
      var freshToken = await refreshOAuthToken(id);
      if (!freshToken) {
        // Refresh failed, fall back to stored token (may still work if recent)
        var tokenData = await browser.storage.local.get([acctKey(id, "oauthToken")]);
        freshToken = tokenData[acctKey(id, "oauthToken")];
      }
      if (!freshToken) {
        console.log("[ExCalSync] No OAuth token for", id);
        return { success: false, reason: "no_token" };
      }
      config.credentials = {
        type: "oauth",
        token: freshToken,
      };
    } else {
      var credData = await browser.storage.local.get([
        acctKey(id, "username"),
        acctKey(id, "password"),
      ]);
      config.credentials = {
        type: "basic",
        username: credData[acctKey(id, "username")] || account.email,
        password: credData[acctKey(id, "password")],
      };
      if (!config.credentials.password) {
        console.log("[ExCalSync] No password for", id);
        return { success: false, reason: "no_password" };
      }
    }

    // Load saved sync state for this account
    var stateData = await browser.storage.local.get([
      acctKey(id, "syncState"),
      acctKey(id, "itemMap"),
      acctKey(id, "reverseMap"),
      acctKey(id, "calendarId"),
    ]);

    var engine = new SyncEngine();
    // Pre-set saved state before init
    engine.syncState = stateData[acctKey(id, "syncState")] || null;
    engine.itemMap = stateData[acctKey(id, "itemMap")] || {};
    engine.reverseMap = stateData[acctKey(id, "reverseMap")] || {};
    engine.calendarId = stateData[acctKey(id, "calendarId")] || null;

    // Override _saveState to use account-scoped keys
    engine._saveState = async function () {
      var stateObj = {};
      stateObj[acctKey(id, "syncState")] = this.syncState;
      stateObj[acctKey(id, "itemMap")] = this.itemMap;
      stateObj[acctKey(id, "reverseMap")] = this.reverseMap;
      stateObj[acctKey(id, "calendarId")] = this.calendarId;
      await browser.storage.local.set(stateObj);
    };

    await engine.init(config);

    // Wire up token refresh so the EWS client can auto-refresh on 401
    if (account.authType === "oauth" && engine.ewsClient) {
      engine.ewsClient._refreshTokenFn = async function () {
        return await refreshOAuthToken(id);
      };
    }

    // Save the discovered EWS URL back to account
    if (!account.ewsUrl && engine.ewsClient && engine.ewsClient.ewsUrl) {
      var acctList = await loadAccounts();
      var acct = acctList.find(function (a) { return a.id === id; });
      if (acct) {
        acct.ewsUrl = engine.ewsClient.ewsUrl;
        await saveAccounts(acctList);
      }
    }

    // Save calendarId
    var calObj = {};
    calObj[acctKey(id, "calendarId")] = engine.calendarId;
    await browser.storage.local.set(calObj);

    // Persist account metadata to calendar properties (survives reinstall)
    await saveAccountToCalendar(engine.calendarId, account);

    // Schedule periodic sync
    await browser.alarms.create(ALARM_PREFIX + id, {
      periodInMinutes: config.syncIntervalMinutes,
    });

    // Do initial full sync
    var result = await engine.fullSync();

    accounts[id] = { engine: engine, lastResult: result, lastError: null };
    console.log("[ExCalSync] Account", id, "sync started, items:", result.total || result.created);
    return { success: true, result: result };
  } catch (e) {
    accounts[id] = { engine: null, lastResult: null, lastError: e.message };
    console.error("[ExCalSync] Account", id, "sync failed:", e);
    return { success: false, error: e.message };
  }
}

/**
 * Stop sync for a single account.
 */
async function stopAccountSync(accountId) {
  await browser.alarms.clear(ALARM_PREFIX + accountId);
  if (accounts[accountId]) {
    accounts[accountId].engine = null;
  }
  console.log("[ExCalSync] Stopped sync for", accountId);
}

/**
 * Handle periodic sync alarms.
 */
browser.alarms.onAlarm.addListener(async function (alarm) {
  if (!alarm.name.startsWith(ALARM_PREFIX)) return;
  var accountId = alarm.name.substring(ALARM_PREFIX.length);
  var acct = accounts[accountId];
  if (!acct || !acct.engine) return;

  try {
    console.log("[ExCalSync] Periodic sync for", accountId);
    acct.lastResult = await acct.engine.incrementalSync();
    acct.lastError = null;
  } catch (e) {
    acct.lastError = e.message;
    console.error("[ExCalSync] Periodic sync failed for", accountId, ":", e);
    if (e.message.includes("401") || e.message.includes("Authentication")) {
      // Try refreshing the OAuth token and retrying once
      var acctList = await loadAccounts();
      var acctInfo = acctList.find(function (a) { return a.id === accountId; });
      if (acctInfo && acctInfo.authType === "oauth") {
        console.log("[ExCalSync] Attempting token refresh for", accountId);
        var newToken = await refreshOAuthToken(accountId);
        if (newToken && acct.engine && acct.engine.ewsClient) {
          acct.engine.ewsClient.credentials.token = newToken;
          try {
            acct.lastResult = await acct.engine.incrementalSync();
            acct.lastError = null;
            console.log("[ExCalSync] Retry after token refresh succeeded for", accountId);
            return;
          } catch (retryErr) {
            acct.lastError = retryErr.message;
            console.error("[ExCalSync] Retry after token refresh also failed:", retryErr.message);
          }
        }
      }
      await stopAccountSync(accountId);
    }
  }
});

/**
 * Handle toolbar button click.
 */
browser.browserAction.onClicked.addListener(async function () {
  browser.runtime.openOptionsPage();
});

/**
 * Handle messages from options page.
 */
browser.runtime.onMessage.addListener(async function (message) {
  switch (message.action) {

    case "getAccounts": {
      var acctList = await loadAccounts();
      var result = [];
      for (var a of acctList) {
        var state = accounts[a.id];
        result.push({
          id: a.id,
          email: a.email,
          authType: a.authType,
          calendarName: a.calendarName,
          ewsUrl: a.ewsUrl,
          syncIntervalMinutes: a.syncIntervalMinutes,
          connected: !!(state && state.engine),
          lastResult: state ? state.lastResult : null,
          lastError: state ? state.lastError : null,
          status: state && state.engine ? state.engine.getStatus() : null,
        });
      }
      return result;
    }

    case "startDeviceCode": {
      try {
        var deviceResponse = await fetch(
          "https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode",
          {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: "client_id=" + encodeURIComponent(MS_CLIENT_ID) +
                  "&scope=" + encodeURIComponent(EWS_SCOPE),
          }
        );

        if (!deviceResponse.ok) {
          return { success: false, error: "Failed to start sign-in: " + deviceResponse.status };
        }

        var deviceData = await deviceResponse.json();
        return {
          success: true,
          userCode: deviceData.user_code,
          deviceCode: deviceData.device_code,
          verificationUri: deviceData.verification_uri || "https://microsoft.com/devicelogin",
          expiresIn: deviceData.expires_in || 900,
          interval: deviceData.interval || 5,
        };
      } catch (e) {
        return { success: false, error: e.message };
      }
    }

    case "pollDeviceCode": {
      var pollInterval = (message.interval || 5) * 1000;
      var expiresAt = Date.now() + (message.expiresIn || 900) * 1000;

      try {
        while (Date.now() < expiresAt) {
          await new Promise(function (r) { setTimeout(r, pollInterval); });

          var tokenResponse = await fetch(
            "https://login.microsoftonline.com/organizations/oauth2/v2.0/token",
            {
              method: "POST",
              headers: { "Content-Type": "application/x-www-form-urlencoded" },
              body: "client_id=" + encodeURIComponent(MS_CLIENT_ID) +
                    "&grant_type=urn:ietf:params:oauth:grant-type:device_code" +
                    "&device_code=" + encodeURIComponent(message.deviceCode),
            }
          );

          if (tokenResponse.ok) {
            var tokenData = await tokenResponse.json();

            // Store token scoped to the account being set up
            var tokenStore = {};
            tokenStore[acctKey(message.accountId, "oauthToken")] = tokenData.access_token;
            tokenStore[acctKey(message.accountId, "oauthRefreshToken")] = tokenData.refresh_token || "";
            await browser.storage.local.set(tokenStore);

            return { success: true };
          }

          var errorBody = await tokenResponse.json().catch(function () { return {}; });
          var errorCode = errorBody.error || "";

          if (errorCode === "authorization_pending") continue;
          else if (errorCode === "slow_down") { pollInterval += 5000; continue; }
          else if (errorCode === "authorization_declined") return { success: false, error: "Sign-in declined." };
          else if (errorCode === "expired_token") return { success: false, error: "Sign-in timed out." };
          else return { success: false, error: errorBody.error_description || errorCode };
        }
        return { success: false, error: "Sign-in timed out." };
      } catch (e) {
        return { success: false, error: e.message };
      }
    }

    case "addAccount": {
      var acctList = await loadAccounts();
      var newId = generateId();
      var newAccount = {
        id: newId,
        email: message.email,
        authType: message.authType || "oauth",
        calendarName: message.calendarName || "Exchange (" + message.email + ")",
        ewsUrl: message.ewsUrl || "",
        syncIntervalMinutes: message.syncIntervalMinutes || DEFAULT_SYNC_INTERVAL_MINUTES,
      };

      if (message.authType === "basic") {
        var credStore = {};
        credStore[acctKey(newId, "username")] = message.username || message.email;
        credStore[acctKey(newId, "password")] = message.password;
        await browser.storage.local.set(credStore);
      } else if (message.authType === "oauth" && message.pendingAccountId) {
        // Move token from pending temp ID to the real account ID
        var pendingTokenKey = acctKey(message.pendingAccountId, "oauthToken");
        var pendingRefreshKey = acctKey(message.pendingAccountId, "oauthRefreshToken");
        var pendingData = await browser.storage.local.get([pendingTokenKey, pendingRefreshKey]);

        var moveStore = {};
        moveStore[acctKey(newId, "oauthToken")] = pendingData[pendingTokenKey] || "";
        moveStore[acctKey(newId, "oauthRefreshToken")] = pendingData[pendingRefreshKey] || "";
        await browser.storage.local.set(moveStore);
        await browser.storage.local.remove([pendingTokenKey, pendingRefreshKey]);
      }

      acctList.push(newAccount);
      await saveAccounts(acctList);

      // Start sync for this account
      var syncResult = await startAccountSync(newAccount);
      return { success: syncResult.success, accountId: newId, error: syncResult.error };
    }

    case "removeAccount": {
      await stopAccountSync(message.accountId);

      // Remove calendar if requested
      var acctState = accounts[message.accountId];
      if (message.removeCalendar && acctState && acctState.engine && acctState.engine.calendarId) {
        try {
          await browser.CalendarAPI.removeCalendar(acctState.engine.calendarId);
        } catch (e) {
          console.error("[ExCalSync] Failed to remove calendar:", e);
        }
      }

      // Clean up storage
      var keysToRemove = [
        acctKey(message.accountId, "oauthToken"),
        acctKey(message.accountId, "oauthRefreshToken"),
        acctKey(message.accountId, "username"),
        acctKey(message.accountId, "password"),
        acctKey(message.accountId, "syncState"),
        acctKey(message.accountId, "itemMap"),
        acctKey(message.accountId, "reverseMap"),
        acctKey(message.accountId, "calendarId"),
      ];
      await browser.storage.local.remove(keysToRemove);

      // Remove from account list
      var acctList = await loadAccounts();
      acctList = acctList.filter(function (a) { return a.id !== message.accountId; });
      await saveAccounts(acctList);

      delete accounts[message.accountId];
      return { success: true };
    }

    case "syncAccount": {
      var acctState = accounts[message.accountId];
      if (!acctState || !acctState.engine) {
        // Try to reconnect
        var acctList = await loadAccounts();
        var acct = acctList.find(function (a) { return a.id === message.accountId; });
        if (acct) {
          console.log("[ExCalSync] syncAccount: engine not connected, attempting reconnect for", message.accountId);
          var reconnResult = await startAccountSync(acct);
          if (reconnResult.success) {
            return { success: true, result: accounts[message.accountId].lastResult, reconnected: true };
          }
          return { success: false, error: "Reconnect failed: " + (reconnResult.error || reconnResult.reason || "unknown") };
        }
        return { success: false, error: "Account not found" };
      }
      try {
        acctState.lastResult = await acctState.engine.incrementalSync();
        acctState.lastError = null;
        return { success: true, result: acctState.lastResult };
      } catch (e) {
        acctState.lastError = e.message;
        return { success: false, error: e.message };
      }
    }

    case "fullSyncAccount": {
      var acctState = accounts[message.accountId];
      if (!acctState || !acctState.engine) {
        // Try to reconnect
        var acctList = await loadAccounts();
        var acct = acctList.find(function (a) { return a.id === message.accountId; });
        if (acct) {
          console.log("[ExCalSync] fullSyncAccount: engine not connected, attempting reconnect for", message.accountId);
          var reconnResult = await startAccountSync(acct);
          if (!reconnResult.success) {
            return { success: false, error: "Reconnect failed: " + (reconnResult.error || reconnResult.reason || "unknown") };
          }
          // startAccountSync already does a fullSync, so we're done
          return { success: true, result: accounts[message.accountId].lastResult, reconnected: true };
        }
        return { success: false, error: "Account not found" };
      }
      try {
        await acctState.engine.resetSync();
        acctState.lastResult = await acctState.engine.fullSync();
        acctState.lastError = null;
        return { success: true, result: acctState.lastResult };
      } catch (e) {
        acctState.lastError = e.message;
        return { success: false, error: e.message };
      }
    }

    default:
      if (message.action === "getPendingInvites") {
        return await loadPendingInvites();
      }
      if (message.action === "acceptInvite") {
        await acceptPendingInvite(message.notifId, message.calendarId);
        return { success: true };
      }
      if (message.action === "dismissInvite") {
        var pending = await loadPendingInvites();
        delete pending[message.notifId];
        await savePendingInvites(pending);
        return { success: true };
      }
      if (message.action === "getAllCalendars") {
        return await browser.CalendarAPI.getCalendars();
      }
      if (message.action === "getMailAccounts") {
        try {
          var mailAccounts = await browser.accounts.list();
          return mailAccounts.map(function (a) {
            var email = "";
            if (a.identities && a.identities.length > 0) {
              email = a.identities[0].email || "";
            }
            return { id: a.id, name: a.name || email, email: email, type: a.type };
          }).filter(function (a) { return a.email; });
        } catch (e) {
          console.error("[ExCalSync] getMailAccounts error:", e);
          return [];
        }
      }
      if (message.action === "getImipAccountMap") {
        var mapStore = await browser.storage.local.get([IMIP_ACCOUNT_MAP_KEY]);
        return mapStore[IMIP_ACCOUNT_MAP_KEY] || {};
      }
      if (message.action === "setImipAccountMap") {
        var mapObj = {};
        mapObj[IMIP_ACCOUNT_MAP_KEY] = message.accountMap || {};
        await browser.storage.local.set(mapObj);
        console.log("[ExCalSync] iMIP account map saved:", JSON.stringify(message.accountMap));
        return { success: true };
      }
      if (message.action === "checkMeetingMessage") {
        return await checkMeetingMessage(message.messageId, message.subject, message.author);
      }
      if (message.action === "respondToMeeting") {
        return await respondToMeetingAction(message.responseType, message.itemId, message.changeKey, message.accountId, message.messageId);
      }
      return { error: "Unknown action: " + message.action };
  }
});

/**
 * =============================================
 * Meeting Response (Accept/Tentative/Decline)
 * =============================================
 * Detects meeting invite emails and responds via EWS.
 */

/**
 * Check if a displayed message is a meeting invite, and find the corresponding
 * Exchange calendar item to enable Accept/Tentative/Decline responses.
 */
async function checkMeetingMessage(messageId, subject, author) {
  try {
    // First: check if message has a text/calendar part (definitive meeting invite)
    var isMeetingByContent = false;
    var parsedInvite = null;

    try {
      var inlineParts = await browser.messages.listInlineTextParts(messageId);
      for (var part of inlineParts) {
        if (isCalendarInvite(part.contentType)) {
          isMeetingByContent = true;
          parsedInvite = parseICalendar(part.content);
          break;
        }
      }
    } catch (e) { /* may fail for some messages */ }

    if (!isMeetingByContent) {
      try {
        var attachments = await browser.messages.listAttachments(messageId);
        for (var att of attachments) {
          if (isCalendarInvite(att.contentType) || (att.name && att.name.toLowerCase().endsWith(".ics"))) {
            isMeetingByContent = true;
            var file = await browser.messages.getAttachmentFile(messageId, att.partName);
            var content = await file.text();
            parsedInvite = parseICalendar(content);
            break;
          }
        }
      } catch (e) { /* may fail */ }
    }

    // Second: check body for OWA calendar links (Exchange sends these instead of .ics)
    var hasOwaCalendarLink = false;
    if (!isMeetingByContent) {
      try {
        var inlineTexts = await browser.messages.listInlineTextParts(messageId);
        for (var tp of inlineTexts) {
          if (tp.content && tp.content.indexOf("calendar/item") !== -1) {
            hasOwaCalendarLink = true;
            break;
          }
          if (tp.content && tp.content.indexOf("itemid=") !== -1 && tp.content.indexOf("outlook.office365.com") !== -1) {
            hasOwaCalendarLink = true;
            break;
          }
        }
      } catch (e) { /* may fail */ }
    }

    // Third: check subject patterns common for meeting invites
    var meetingSubjectPatterns = /^(accepted|tentative|declined|cancelled|updated|invitation):/i;
    var hasMeetingSubject = subject && meetingSubjectPatterns.test(subject.trim());

    if (!isMeetingByContent && !hasOwaCalendarLink && !hasMeetingSubject) {
      return { isMeeting: false };
    }

    // Find the corresponding Exchange calendar item
    var acctList = await loadAccounts();
    for (var acct of acctList) {
      var acctState = accounts[acct.id];
      if (!acctState || !acctState.engine || !acctState.engine.ewsClient) continue;

      // Search by subject in a wide date range
      var searchSubject = subject || "";
      // Strip common prefixes
      searchSubject = searchSubject.replace(/^(Re:|Fwd:|Accepted:|Tentative:|Declined:|Cancelled:|Updated:|Invitation:)\s*/gi, "").trim();

      if (!searchSubject) continue;

      try {
        var now = new Date();
        var searchStart = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000); // 1 week ago
        var searchEnd = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000); // 90 days forward

        var found = await acctState.engine.ewsClient.findItemBySubject(searchSubject, searchStart, searchEnd);

        if (found.length > 0) {
          var item = found[0]; // Best match (subject search)
          return {
            isMeeting: true,
            subject: item.subject || searchSubject,
            start: item.start || (parsedInvite ? parsedInvite.dtstart : null),
            end: item.end || (parsedInvite ? parsedInvite.dtend : null),
            location: item.location || (parsedInvite ? parsedInvite.location : null),
            organizer: item.organizer || (parsedInvite ? (parsedInvite.organizerName || parsedInvite.organizer) : null) || author,
            itemId: item.id,
            changeKey: item.changeKey,
            accountId: acct.id,
            messageId: messageId,
          };
        }
      } catch (e) {
        console.error("[ExCalSync] Meeting search failed for account", acct.id, ":", e.message);
      }
    }

    // If we know it's a meeting but can't find the item (maybe not synced yet)
    if (isMeetingByContent || hasOwaCalendarLink) {
      return {
        isMeeting: true,
        subject: (parsedInvite ? parsedInvite.summary : null) || subject || "Meeting",
        start: parsedInvite ? parsedInvite.dtstart : null,
        end: parsedInvite ? parsedInvite.dtend : null,
        location: parsedInvite ? parsedInvite.location : null,
        organizer: parsedInvite ? (parsedInvite.organizerName || parsedInvite.organizer) : author,
        itemId: null,
        changeKey: null,
        accountId: null,
        messageId: messageId,
      };
    }

    return { isMeeting: false };
  } catch (e) {
    console.error("[ExCalSync] checkMeetingMessage error:", e);
    return { isMeeting: false };
  }
}

/**
 * Send a meeting response (Accept/Tentative/Decline) via EWS.
 */
async function respondToMeetingAction(responseType, itemId, changeKey, accountId, messageId) {
  try {
    if (!itemId || !changeKey) {
      // Try to find the item first by triggering a sync
      if (accountId) {
        var acctState = accounts[accountId];
        if (acctState && acctState.engine) {
          await acctState.engine.incrementalSync();
        }
      }
      return { success: false, error: "Meeting not found on Exchange. Try syncing first, then respond again." };
    }

    if (!accountId) {
      return { success: false, error: "No Exchange account found for this meeting." };
    }

    var acctState = accounts[accountId];
    if (!acctState || !acctState.engine || !acctState.engine.ewsClient) {
      return { success: false, error: "Exchange account not connected." };
    }

    // Send the response
    await acctState.engine.ewsClient.respondToMeeting(itemId, changeKey, responseType);
    console.log("[ExCalSync] Meeting response sent:", responseType, "for item:", itemId.substring(0, 20) + "...");

    // Trigger a sync to update local state
    try {
      await acctState.engine.incrementalSync();
    } catch (e) { /* non-critical */ }

    // Mark the email as read
    if (messageId) {
      try {
        await browser.messages.update(messageId, { read: true });
      } catch (e) { /* may fail */ }
    }

    return { success: true };
  } catch (e) {
    console.error("[ExCalSync] respondToMeeting error:", e);
    return { success: false, error: e.message };
  }
}

/**
 * =============================================
 * Meeting Request (iMIP) Processing
 * =============================================
 * Monitors incoming emails for text/calendar parts (meeting invitations),
 * parses them, and adds/updates/removes events in the appropriate calendar.
 */

var MEETING_UID_MAP_KEY = "meetingUidMap";
var PENDING_INVITES_KEY = "pendingInvites";
var IMIP_ACCOUNT_MAP_KEY = "imipAccountCalendarMap";

/**
 * Load the meeting UID map from storage.
 * Maps iCalendar UID -> { calendarId, localItemId, sequence }
 */
async function loadMeetingUidMap() {
  var data = await browser.storage.local.get([MEETING_UID_MAP_KEY]);
  return data[MEETING_UID_MAP_KEY] || {};
}

async function saveMeetingUidMap(map) {
  var obj = {};
  obj[MEETING_UID_MAP_KEY] = map;
  await browser.storage.local.set(obj);
}

/**
 * Load/save pending invitations (for notification-based accept flow).
 */
async function loadPendingInvites() {
  var data = await browser.storage.local.get([PENDING_INVITES_KEY]);
  return data[PENDING_INVITES_KEY] || {};
}

async function savePendingInvites(invites) {
  var obj = {};
  obj[PENDING_INVITES_KEY] = invites;
  await browser.storage.local.set(obj);
}

/**
 * Find target calendars for a meeting based on attendee emails.
 * Returns an array of { calendarId, accountId, isExchange }.
 * Matches our accounts against the ATTENDEE list, skipping the organizer.
 */
async function findTargetCalendars(parsed, receivingAccountId) {
  var acctList = await loadAccounts();
  var targets = [];

  // Build set of attendee emails
  var attendeeEmails = new Set();
  if (parsed.attendees) {
    for (var att of parsed.attendees) {
      if (att.email) attendeeEmails.add(att.email.toLowerCase());
    }
  }

  var organizerEmail = parsed.organizer ? parsed.organizer.toLowerCase() : "";

  // 1) Match our accounts that appear as attendees (not organizer)
  for (var acct of acctList) {
    var acctEmail = acct.email.toLowerCase();
    if (acctEmail === organizerEmail) continue; // organizer already has it via sync
    if (attendeeEmails.has(acctEmail)) {
      var state = accounts[acct.id];
      if (state && state.engine && state.engine.calendarId) {
        targets.push({ calendarId: state.engine.calendarId, accountId: acct.id, isExchange: true });
      }
    }
  }
  if (targets.length > 0) return targets;

  // 2) Fallback: domain match, but still skip organizer's own account
  var orgDomain = imipGetDomain(organizerEmail);
  for (var acct of acctList) {
    var acctEmail = acct.email.toLowerCase();
    if (acctEmail === organizerEmail) continue;
    var acctDomain = imipGetDomain(acctEmail);
    if (acctDomain && orgDomain && acctDomain === orgDomain) {
      var state = accounts[acct.id];
      if (state && state.engine && state.engine.calendarId) {
        targets.push({ calendarId: state.engine.calendarId, accountId: acct.id, isExchange: true });
      }
    }
  }
  if (targets.length > 0) return targets;

  // 3) Single account fallback (only if it's not the organizer)
  if (acctList.length === 1) {
    var singleAcct = acctList[0];
    if (singleAcct.email.toLowerCase() !== organizerEmail) {
      var singleState = accounts[singleAcct.id];
      if (singleState && singleState.engine && singleState.engine.calendarId) {
        return [{ calendarId: singleState.engine.calendarId, accountId: singleAcct.id, isExchange: true }];
      }
    }
  }

  // 4) No Exchange match — look up per-account calendar mapping
  var mapData = await browser.storage.local.get([IMIP_ACCOUNT_MAP_KEY]);
  var accountMap = mapData[IMIP_ACCOUNT_MAP_KEY] || {};

  if (receivingAccountId && accountMap[receivingAccountId]) {
    var mappedCalId = accountMap[receivingAccountId];
    if (mappedCalId === "disabled") {
      console.log("[ExCalSync iMIP] iMIP disabled for mail account", receivingAccountId);
      return [];
    }
    // Verify the mapped calendar still exists
    try {
      var calendars = await browser.CalendarAPI.getCalendars();
      var exists = calendars.some(function (c) { return c.id === mappedCalId; });
      if (exists) {
        console.log("[ExCalSync iMIP] Using mapped calendar for account", receivingAccountId);
        return [{ calendarId: mappedCalId, accountId: null, isExchange: false }];
      }
    } catch (e) { /* fall through */ }
    console.log("[ExCalSync iMIP] Mapped calendar not found for account", receivingAccountId);
  }

  // No mapping configured for this account — let Thunderbird handle it
  console.log("[ExCalSync iMIP] No iMIP mapping for mail account", receivingAccountId || "(unknown)", ", skipping");
  return [];
}

/**
 * Check if a meeting UID already exists in any synced account's item map.
 * EWS-synced items use UIDs like "ews-<hash>@exchange-cal-sync".
 */
function isUidInSyncMap(uid) {
  // The UID in the iCalendar invite is the original Exchange/Outlook UID.
  // Our EWS sync generates synthetic UIDs from Exchange ItemIds.
  // We can't match by UID directly, so we rely on the meetingUidMap instead.
  return false;
}

/**
 * Handle a parsed meeting REQUEST.
 */
async function handleMeetingRequest(parsed, messageId, target) {
  console.log("[ExCalSync iMIP] Meeting REQUEST:", parsed.summary, "UID:", parsed.uid,
    "-> account:", target.accountId || "fallback");

  if (target.isExchange) {
    // For Exchange-matched accounts, trigger an incremental sync
    // The Exchange server will already have the meeting invitation
    var acctState = accounts[target.accountId];
    if (acctState && acctState.engine) {
      console.log("[ExCalSync iMIP] Exchange account matched, triggering sync for account:", target.accountId);
      try {
        await acctState.engine.incrementalSync();
        await browser.notifications.create("imip-synced-" + Date.now(), {
          type: "basic",
          title: "Meeting synced",
          message: (parsed.summary || "Meeting") + "\n" + (parsed.dtstart || ""),
          iconUrl: "icons/icon-64.svg",
        });
      } catch (e) {
        console.error("[ExCalSync iMIP] Sync after meeting request failed:", e);
      }
    }
    return;
  }

  // Non-Exchange calendar — check sequence and prompt user
  var uidMap = await loadMeetingUidMap();
  var existing = uidMap[parsed.uid];
  if (existing && existing.sequence >= parsed.sequence) {
    console.log("[ExCalSync iMIP] Already have this meeting (seq " + existing.sequence + "), skipping");
    return;
  }

  var notifId = "imip-request-" + Date.now();
  var timeStr = parsed.dtstart || "";
  if (parsed.dtend) timeStr += " - " + parsed.dtend;

  var pendingInvites = await loadPendingInvites();
  pendingInvites[notifId] = {
    parsed: parsed,
    calendarId: target.calendarId,
    messageId: messageId,
    timestamp: Date.now(),
  };
  await savePendingInvites(pendingInvites);

  await browser.notifications.create(notifId, {
    type: "basic",
    title: "Meeting Invitation: " + (parsed.summary || "No Subject"),
    message: (parsed.organizerName || parsed.organizer || "Unknown") + "\n" +
             timeStr + (parsed.location ? "\n" + parsed.location : "") +
             "\n\nClick to accept.",
    iconUrl: "icons/icon-64.svg",
  });
}

/**
 * Handle a parsed meeting CANCEL.
 */
async function handleMeetingCancel(parsed, messageId, target) {
  console.log("[ExCalSync iMIP] Meeting CANCEL:", parsed.summary, "UID:", parsed.uid,
    "-> account:", target.accountId || "fallback");

  if (target.isExchange) {
    var acctState = accounts[target.accountId];
    if (acctState && acctState.engine) {
      console.log("[ExCalSync iMIP] Exchange account matched for cancel, triggering sync");
      try {
        await acctState.engine.incrementalSync();
      } catch (e) {
        console.error("[ExCalSync iMIP] Sync after cancel failed:", e);
      }
    }
    await browser.notifications.create("imip-cancel-" + Date.now(), {
      type: "basic",
      title: "Meeting cancelled",
      message: (parsed.summary || "Meeting") + " has been cancelled.",
      iconUrl: "icons/icon-64.svg",
    });
    return;
  }

  // Non-Exchange: look up in meetingUidMap and delete
  var uidMap = await loadMeetingUidMap();
  var existing = uidMap[parsed.uid];
  if (existing) {
    try {
      await browser.CalendarAPI.deleteItem(existing.calendarId, existing.localItemId);
      console.log("[ExCalSync iMIP] Deleted cancelled meeting from calendar");
    } catch (e) {
      console.error("[ExCalSync iMIP] Failed to delete cancelled meeting:", e);
    }
    delete uidMap[parsed.uid];
    await saveMeetingUidMap(uidMap);
  }

  await browser.notifications.create("imip-cancel-" + Date.now(), {
    type: "basic",
    title: "Meeting cancelled",
    message: (parsed.summary || "Meeting") + " has been cancelled.",
    iconUrl: "icons/icon-64.svg",
  });
}

/**
 * Handle a meeting update (REQUEST with higher sequence).
 */
async function handleMeetingUpdate(parsed, messageId, target) {
  console.log("[ExCalSync iMIP] Meeting UPDATE:", parsed.summary, "UID:", parsed.uid, "seq:", parsed.sequence,
    "-> account:", target.accountId || "fallback");

  if (target.isExchange) {
    var acctState = accounts[target.accountId];
    if (acctState && acctState.engine) {
      try { await acctState.engine.incrementalSync(); } catch (e) { /* logged elsewhere */ }
    }
    return;
  }

  var uidMap = await loadMeetingUidMap();
  var existing = uidMap[parsed.uid];

  if (existing) {
    try {
      await browser.CalendarAPI.updateItem(existing.calendarId, existing.localItemId, parsed.rawIcs);
      uidMap[parsed.uid].sequence = parsed.sequence;
      await saveMeetingUidMap(uidMap);
      console.log("[ExCalSync iMIP] Updated meeting in calendar");

      await browser.notifications.create("imip-update-" + Date.now(), {
        type: "basic",
        title: "Meeting updated",
        message: (parsed.summary || "Meeting") + "\n" + (parsed.dtstart || ""),
        iconUrl: "icons/icon-64.svg",
      });
    } catch (e) {
      console.error("[ExCalSync iMIP] Failed to update meeting:", e);
    }
  } else {
    await handleMeetingRequest(parsed, messageId, target);
  }
}

/**
 * Accept a pending meeting invitation (called from notification click or options page).
 */
async function acceptPendingInvite(notifId, overrideCalendarId) {
  var pendingInvites = await loadPendingInvites();
  var invite = pendingInvites[notifId];
  if (!invite) {
    console.log("[ExCalSync iMIP] No pending invite found for:", notifId);
    return;
  }

  var targetCalId = overrideCalendarId || invite.calendarId;

  try {
    var result = await browser.CalendarAPI.addItem(targetCalId, invite.parsed.rawIcs);
    console.log("[ExCalSync iMIP] Added meeting to calendar", targetCalId, "localId:", result.id);

    // Update UID map
    var uidMap = await loadMeetingUidMap();
    uidMap[invite.parsed.uid] = {
      calendarId: targetCalId,
      localItemId: result.id,
      sequence: invite.parsed.sequence,
    };
    await saveMeetingUidMap(uidMap);

    // Mark email as read
    if (invite.messageId) {
      try {
        await browser.messages.update(invite.messageId, { read: true });
      } catch (e) { /* message may have been moved/deleted */ }
    }
  } catch (e) {
    console.error("[ExCalSync iMIP] Failed to accept invitation:", e);
  }

  delete pendingInvites[notifId];
  await savePendingInvites(pendingInvites);
}

/**
 * Process a single message for iCalendar content.
 */
async function processMessageForInvites(messageId, receivingAccountId) {
  try {
    // Check inline text parts for text/calendar
    var inlineParts = await browser.messages.listInlineTextParts(messageId);
    for (var part of inlineParts) {
      if (isCalendarInvite(part.contentType)) {
        var parsed = parseICalendar(part.content);
        if (parsed) {
          await routeMeetingAction(parsed, messageId, receivingAccountId);
          return; // One invite per message is typical
        }
      }
    }

    // Check attachments for .ics files
    var attachments = await browser.messages.listAttachments(messageId);
    for (var att of attachments) {
      if (isCalendarInvite(att.contentType) || (att.name && att.name.toLowerCase().endsWith(".ics"))) {
        var file = await browser.messages.getAttachmentFile(messageId, att.partName);
        var content = await file.text();
        var parsed = parseICalendar(content);
        if (parsed) {
          await routeMeetingAction(parsed, messageId, receivingAccountId);
          return;
        }
      }
    }
  } catch (e) {
    console.error("[ExCalSync iMIP] Error processing message", messageId, ":", e);
  }
}

/**
 * Route a parsed iCalendar to the appropriate handler based on METHOD.
 */
async function routeMeetingAction(parsed, messageId, receivingAccountId) {
  var targets = await findTargetCalendars(parsed, receivingAccountId);
  if (targets.length === 0) {
    console.log("[ExCalSync iMIP] No target calendar found, skipping");
    return;
  }

  console.log("[ExCalSync iMIP] Routing", parsed.method, "to", targets.length, "target(s)");

  for (var target of targets) {
    switch (parsed.method) {
      case "REQUEST": {
        var uidMap = await loadMeetingUidMap();
        var existing = uidMap[parsed.uid];
        if (existing && parsed.sequence > existing.sequence) {
          await handleMeetingUpdate(parsed, messageId, target);
        } else {
          await handleMeetingRequest(parsed, messageId, target);
        }
        break;
      }
      case "CANCEL":
        await handleMeetingCancel(parsed, messageId, target);
        break;
      default:
        console.log("[ExCalSync iMIP] Ignoring method:", parsed.method);
        break;
    }
  }
}

/**
 * Listen for new incoming emails and scan for meeting invitations.
 */
browser.messages.onNewMailReceived.addListener(async function (folder, messages) {
  console.log("[ExCalSync iMIP] New mail received:", messages.messages.length, "messages in", folder.name);
  var receivingAccountId = folder.accountId || null;
  for (var msg of messages.messages) {
    await processMessageForInvites(msg.id, receivingAccountId);
  }
});

/**
 * Handle notification clicks — accept pending invitations.
 */
browser.notifications.onClicked.addListener(async function (notifId) {
  if (notifId.indexOf("imip-request-") === 0) {
    await acceptPendingInvite(notifId);
    await browser.notifications.clear(notifId);
  }
});

/**
 * Auto-start all configured accounts on load.
 * If storage is empty (e.g. after reinstall), recover accounts from calendar properties.
 */
(async function () {
  try {
    // Migrate old single-account format if present
    var oldData = await browser.storage.local.get(["configured", "email", "accounts"]);
    if (oldData.configured && oldData.email && !oldData.accounts) {
      console.log("[ExCalSync] Migrating old single-account to multi-account...");
      var oldFull = await browser.storage.local.get(null);
      var migratedId = generateId();
      var migrated = {
        id: migratedId,
        email: oldFull.email,
        authType: oldFull.authType || "basic",
        calendarName: oldFull.calendarName || "Exchange Calendar",
        ewsUrl: oldFull.ewsUrl || "",
        syncIntervalMinutes: oldFull.syncIntervalMinutes || DEFAULT_SYNC_INTERVAL_MINUTES,
      };

      // Copy credentials to scoped keys
      var migrateStore = {};
      if (oldFull.oauthToken) migrateStore[acctKey(migratedId, "oauthToken")] = oldFull.oauthToken;
      if (oldFull.oauthRefreshToken) migrateStore[acctKey(migratedId, "oauthRefreshToken")] = oldFull.oauthRefreshToken;
      if (oldFull.username) migrateStore[acctKey(migratedId, "username")] = oldFull.username;
      if (oldFull.password) migrateStore[acctKey(migratedId, "password")] = oldFull.password;
      if (oldFull.syncState) migrateStore[acctKey(migratedId, "syncState")] = oldFull.syncState;
      if (oldFull.itemMap) migrateStore[acctKey(migratedId, "itemMap")] = oldFull.itemMap;
      if (oldFull.reverseMap) migrateStore[acctKey(migratedId, "reverseMap")] = oldFull.reverseMap;
      if (oldFull.calendarId) migrateStore[acctKey(migratedId, "calendarId")] = oldFull.calendarId;

      migrateStore.accounts = [migrated];
      await browser.storage.local.set(migrateStore);

      // Clean old flat keys
      await browser.storage.local.remove([
        "configured", "email", "username", "password", "ewsUrl",
        "authType", "oauthToken", "oauthRefreshToken",
        "calendarName", "syncIntervalMinutes",
        "syncState", "itemMap", "reverseMap", "calendarId",
      ]);
      console.log("[ExCalSync] Migration complete");
    }

    var acctList = await loadAccounts();

    // Recovery: if storage is empty, try to recover from calendar properties
    if (acctList.length === 0) {
      console.log("[ExCalSync] No accounts in storage, attempting recovery from calendar properties...");
      var recovered = await recoverAccountsFromCalendars();
      if (recovered.length > 0) {
        console.log("[ExCalSync] Recovered", recovered.length, "accounts from calendars");
        await saveAccounts(recovered);
        acctList = recovered;
      }
    }

    console.log("[ExCalSync] Starting", acctList.length, "accounts...");
    for (var acct of acctList) {
      await startAccountSync(acct);
    }
  } catch (e) {
    console.error("[ExCalSync] Auto-start failed:", e);
  }
})();
