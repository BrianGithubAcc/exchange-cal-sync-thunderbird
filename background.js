console.log("[ExCalSync] background.js loading");
/**
 * Background script for Exchange Calendar Sync.
 * Supports multiple Exchange accounts, each with its own sync engine.
 */

// Depends on: sync/engine.js (SyncEngine, DEFAULT_SYNC_INTERVAL_MINUTES)
// Depends on: ews/client.js (EwsClient)

var ALARM_PREFIX = "excalsync-";
// Access tokens are deliberately process-memory-only. Refresh tokens are
// stored by the privileged SecureStorage Experiment in Thunderbird's Login
// Manager rather than browser.storage.local.
var oauthTokens = Object.create(null);
var oauthPollCancelled = Object.create(null);

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

function secureTokenKey(accountId) {
  return "oauth-refresh-" + accountId;
}

function oauthResponseBody(response) {
  if (!response || typeof response.body !== "string") return {};
  try {
    return response.body ? JSON.parse(response.body) : {};
  } catch (e) {
    return { error: "invalid_response" };
  }
}

async function getRefreshToken(accountId) {
  var result = await browser.SecureStorage.get(secureTokenKey(accountId));
  return result && result.found ? result.value : null;
}

async function setRefreshToken(accountId, token) {
  if (token) {
    await browser.SecureStorage.set(secureTokenKey(accountId), token);
  } else {
    await browser.SecureStorage.remove(secureTokenKey(accountId));
  }
}

async function clearOAuthCredentials(accountId) {
  delete oauthTokens[accountId];
  await setRefreshToken(accountId, null);
}

async function migrateLegacyOAuthStorage() {
  const all = await browser.storage.local.get(null);
  const removeKeys = [];
  for (const key of Object.keys(all)) {
    const refreshMatch = key.match(/^acct_(.+)_oauthRefreshToken$/);
    if (refreshMatch && all[key]) {
      await setRefreshToken(refreshMatch[1], all[key]);
      removeKeys.push(key);
      continue;
    }
    if (/^acct_.+_oauthToken$/.test(key)) {
      // Access tokens from older releases are intentionally not migrated.
      removeKeys.push(key);
      continue;
    }
    if (/^acct_.+_(username|password)$/.test(key)) {
      // Remove credentials written by the old Basic Auth implementation.
      removeKeys.push(key);
    }
  }
  if (removeKeys.length > 0) {
    await browser.storage.local.remove(removeKeys);
  }
}

/**
 * Refresh an OAuth access token using the stored refresh token.
 * Returns the new access token, or null if refresh failed.
 */
async function refreshOAuthToken(accountId) {
  var refreshToken = await getRefreshToken(accountId);

  if (!refreshToken) {
    return null;
  }

  try {
    var response = await browser.SecureStorage.refreshAccessToken(refreshToken);

    if (response.status < 200 || response.status >= 300) {
      var errBody = oauthResponseBody(response);
      console.error("[ExCalSync] Token refresh failed:", response.status, errBody.error_description || errBody.error || "");
      return null;
    }

    var tokenData = oauthResponseBody(response);
    if (!tokenData.access_token) return null;
    oauthTokens[accountId] = tokenData.access_token;
    if (tokenData.refresh_token) {
      await setRefreshToken(accountId, tokenData.refresh_token);
    }
    return tokenData.access_token;
  } catch (e) {
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
        // Access tokens are memory-only. A refresh token is required after a
        // Thunderbird restart or when the in-memory token has expired.
        freshToken = oauthTokens[id];
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
      return { success: false, reason: "oauth_required" };
    }

    // Load saved sync state for this account
    var stateData = await browser.storage.local.get([
      acctKey(id, "syncState"),
      acctKey(id, "itemMap"),
      acctKey(id, "reverseMap"),
      acctKey(id, "calendarId"),
      acctKey(id, "localChangesQueue"),
      acctKey(id, "conflicts"),
    ]);

    var engine = new SyncEngine();
    // Pre-set saved state before init
    engine.syncState = stateData[acctKey(id, "syncState")] || null;
    engine.itemMap = stateData[acctKey(id, "itemMap")] || {};
    engine.reverseMap = stateData[acctKey(id, "reverseMap")] || {};
    engine.calendarId = stateData[acctKey(id, "calendarId")] || null;
    engine._localChangesQueue = stateData[acctKey(id, "localChangesQueue")] || [];
    engine.conflicts = stateData[acctKey(id, "conflicts")] || {};

    // Override _saveState to use account-scoped keys
    engine._saveState = async function () {
      var stateObj = {};
      stateObj[acctKey(id, "syncState")] = this.syncState;
      stateObj[acctKey(id, "itemMap")] = this.itemMap;
      stateObj[acctKey(id, "reverseMap")] = this.reverseMap;
      stateObj[acctKey(id, "calendarId")] = this.calendarId;
      stateObj[acctKey(id, "localChangesQueue")] = this._localChangesQueue;
      stateObj[acctKey(id, "conflicts")] = this.conflicts;
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
      if (!message.accountId || typeof message.accountId !== "string") {
        return { success: false, error: "Missing temporary account ID" };
      }
      oauthPollCancelled[message.accountId] = false;
      try {
        var deviceResponse = await browser.SecureStorage.startDeviceCode();

        if (deviceResponse.status < 200 || deviceResponse.status >= 300) {
          var startError = oauthResponseBody(deviceResponse);
          return {
            success: false,
            error: "Failed to start sign-in: " + deviceResponse.status +
              (startError.error_description ? " (" + startError.error_description + ")" : ""),
          };
        }

        var deviceData = oauthResponseBody(deviceResponse);
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
          if (oauthPollCancelled[message.accountId]) {
            await clearOAuthCredentials(message.accountId);
            return { success: false, error: "Sign-in cancelled." };
          }
          await new Promise(function (r) { setTimeout(r, pollInterval); });
          if (oauthPollCancelled[message.accountId]) {
            await clearOAuthCredentials(message.accountId);
            return { success: false, error: "Sign-in cancelled." };
          }

          var tokenResponse = await browser.SecureStorage.pollDeviceCode(message.deviceCode);

          if (tokenResponse.status >= 200 && tokenResponse.status < 300) {
            var tokenData = oauthResponseBody(tokenResponse);

            // Keep access tokens in memory and put only the refresh token in
            // Thunderbird's password manager.
            oauthTokens[message.accountId] = tokenData.access_token;
            if (tokenData.refresh_token) {
              await setRefreshToken(message.accountId, tokenData.refresh_token);
            }

            delete oauthPollCancelled[message.accountId];
            return { success: true };
          }

          var errorBody = oauthResponseBody(tokenResponse);
          var errorCode = errorBody.error || "";

          if (errorCode === "authorization_pending") continue;
          else if (errorCode === "slow_down") { pollInterval += 5000; continue; }
          else if (errorCode === "authorization_declined") {
            await clearOAuthCredentials(message.accountId);
            return { success: false, error: "Sign-in declined." };
          }
          else if (errorCode === "expired_token") {
            await clearOAuthCredentials(message.accountId);
            return { success: false, error: "Sign-in timed out." };
          }
          else return { success: false, error: errorBody.error_description || errorCode };
        }
        await clearOAuthCredentials(message.accountId);
        return { success: false, error: "Sign-in timed out." };
      } catch (e) {
        await clearOAuthCredentials(message.accountId).catch(function () {});
        return { success: false, error: e.message };
      } finally {
        delete oauthPollCancelled[message.accountId];
      }
    }

    case "cancelDeviceCode": {
      if (message.accountId) {
        oauthPollCancelled[message.accountId] = true;
        await clearOAuthCredentials(message.accountId);
      }
      return { success: true };
    }

    case "addAccount": {
      var acctList = await loadAccounts();
      var newId = generateId();
      if (message.authType && message.authType !== "oauth") {
        return { success: false, error: "Only OAuth sign-in is supported." };
      }
      if (!message.pendingAccountId || !oauthTokens[message.pendingAccountId]) {
        return { success: false, error: "Complete Microsoft sign-in first." };
      }
      var newAccount = {
        id: newId,
        email: message.email,
        authType: "oauth",
        calendarName: message.calendarName || "Exchange (" + message.email + ")",
        ewsUrl: "",
        syncIntervalMinutes: message.syncIntervalMinutes || DEFAULT_SYNC_INTERVAL_MINUTES,
      };

      if (message.pendingAccountId) {
        // Move the in-memory access token and password-manager refresh token
        // from the temporary setup ID to the persistent account ID.
        if (oauthTokens[message.pendingAccountId]) {
          oauthTokens[newId] = oauthTokens[message.pendingAccountId];
        }
        var pendingRefresh = await getRefreshToken(message.pendingAccountId);
        if (pendingRefresh) {
          await setRefreshToken(newId, pendingRefresh);
        }
        await clearOAuthCredentials(message.pendingAccountId);
      }

      acctList.push(newAccount);
      await saveAccounts(acctList);

      // Start sync for this account
      var syncResult = await startAccountSync(newAccount);
      return { success: syncResult.success, accountId: newId, error: syncResult.error };
    }

    case "removeAccount": {
      await stopAccountSync(message.accountId);
      await clearOAuthCredentials(message.accountId);

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
        acctKey(message.accountId, "username"),
        acctKey(message.accountId, "password"),
        acctKey(message.accountId, "syncState"),
        acctKey(message.accountId, "itemMap"),
        acctKey(message.accountId, "reverseMap"),
        acctKey(message.accountId, "calendarId"),
        acctKey(message.accountId, "localChangesQueue"),
        acctKey(message.accountId, "conflicts"),
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
      return { error: "Unknown action: " + message.action };
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
        // Legacy password authentication is deliberately not migrated.
        // The user must sign in again with OAuth.
        authType: "oauth",
        calendarName: oldFull.calendarName || "Exchange Calendar",
        ewsUrl: oldFull.ewsUrl || "",
        syncIntervalMinutes: oldFull.syncIntervalMinutes || DEFAULT_SYNC_INTERVAL_MINUTES,
      };

      // Copy non-secret credentials/state to scoped keys. OAuth access tokens
      // remain memory-only; move a legacy refresh token into Login Manager.
      var migrateStore = {};
      if (oldFull.oauthToken) oauthTokens[migratedId] = oldFull.oauthToken;
      if (oldFull.oauthRefreshToken) await setRefreshToken(migratedId, oldFull.oauthRefreshToken);
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

    await migrateLegacyOAuthStorage();

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
