"use strict";

var statusBar = document.getElementById("statusBar");
var accountListEl = document.getElementById("accountList");
var addAccountForm = document.getElementById("addAccountForm");
var btnShowAddAccount = document.getElementById("btnShowAddAccount");
var btnCancelAdd = document.getElementById("btnCancelAdd");
var btnAddAccount = document.getElementById("btnAddAccount");
var btnSignInMs = document.getElementById("btnSignInMs");
var toggleBasicAuth = document.getElementById("toggleBasicAuth");
var basicAuthSection = document.getElementById("basicAuthSection");
var deviceCodeBox = document.getElementById("deviceCodeBox");

var authMode = "oauth";
var oauthReady = false;
var pendingAccountId = null;

function showStatus(msg, type) {
  statusBar.textContent = msg;
  statusBar.className = "status-bar visible " + (type || "info");
}
function hideStatus() { statusBar.className = "status-bar"; }

// Show/hide add form
btnShowAddAccount.addEventListener("click", function () {
  addAccountForm.classList.add("visible");
  btnShowAddAccount.style.display = "none";
  oauthReady = false;
  authMode = "oauth";
  pendingAccountId = Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
});

btnCancelAdd.addEventListener("click", function () {
  addAccountForm.classList.remove("visible");
  btnShowAddAccount.style.display = "";
  deviceCodeBox.classList.remove("visible");
  hideStatus();
});

toggleBasicAuth.addEventListener("click", function (e) {
  e.preventDefault();
  if (basicAuthSection.classList.contains("visible")) {
    basicAuthSection.classList.remove("visible");
    toggleBasicAuth.textContent = "On-premise Exchange? Use password instead";
    authMode = "oauth";
  } else {
    basicAuthSection.classList.add("visible");
    toggleBasicAuth.textContent = "Microsoft 365? Sign in above instead";
    authMode = "basic";
  }
});

// Microsoft sign-in
btnSignInMs.addEventListener("click", async function () {
  var email = document.getElementById("newEmail").value.trim();
  if (!email || email.indexOf("@") === -1) {
    showStatus("Enter your email address first.", "error");
    return;
  }
  btnSignInMs.disabled = true;
  hideStatus();

  try {
    var resp = await browser.runtime.sendMessage({ action: "startDeviceCode" });
    if (!resp.success) {
      showStatus("Sign-in failed: " + resp.error, "error");
      btnSignInMs.disabled = false;
      return;
    }

    document.getElementById("deviceCodeValue").textContent = resp.userCode;
    var urlEl = document.getElementById("deviceCodeUrl");
    urlEl.href = resp.verificationUri;
    urlEl.textContent = resp.verificationUri;
    deviceCodeBox.classList.add("visible");
    document.getElementById("deviceCodeStatus").textContent = "Waiting for sign-in...";
    showStatus("Open the link and enter the code shown.", "info");

    // Generate a temporary account ID for storing the token
    pendingAccountId = Date.now().toString(36) + Math.random().toString(36).substring(2, 8);

    var poll = await browser.runtime.sendMessage({
      action: "pollDeviceCode",
      deviceCode: resp.deviceCode,
      interval: resp.interval || 5,
      expiresIn: resp.expiresIn || 900,
      accountId: pendingAccountId,
    });

    if (poll.success) {
      document.getElementById("deviceCodeStatus").textContent = "Signed in!";
      deviceCodeBox.classList.remove("visible");
      oauthReady = true;
      showStatus("Signed in! Click 'Save & Connect' to add this account.", "success");
    } else {
      document.getElementById("deviceCodeStatus").textContent = "Failed.";
      showStatus("Sign-in failed: " + (poll.error || "Timed out"), "error");
    }
  } catch (e) {
    showStatus("Error: " + e.message, "error");
  } finally {
    btnSignInMs.disabled = false;
  }
});

// Save & Connect
btnAddAccount.addEventListener("click", async function () {
  var email = document.getElementById("newEmail").value.trim();
  if (!email || email.indexOf("@") === -1) {
    showStatus("Enter a valid email address.", "error");
    return;
  }

  if (authMode === "oauth" && !oauthReady) {
    showStatus("Sign in with Microsoft first.", "error");
    return;
  }
  if (authMode === "basic" && !document.getElementById("newPassword").value) {
    showStatus("Enter your password.", "error");
    return;
  }

  btnAddAccount.disabled = true;
  btnAddAccount.textContent = "Connecting...";
  hideStatus();

  try {
    var msg = {
      action: "addAccount",
      email: email,
      authType: authMode,
      calendarName: document.getElementById("newCalName").value.trim() || "Exchange (" + email + ")",
      ewsUrl: document.getElementById("newEwsUrl").value.trim() || "",
      syncIntervalMinutes: parseInt(document.getElementById("newInterval").value, 10) || 5,
    };

    if (authMode === "basic") {
      msg.username = document.getElementById("newUsername").value.trim() || email;
      msg.password = document.getElementById("newPassword").value;
    }

    // For OAuth, we need to rename the temp token keys to match the server-generated ID
    // Actually, the background generates its own ID. We need to move the token.
    // Let's pass the pendingAccountId so background can use the stored token
    if (authMode === "oauth") {
      msg.pendingAccountId = pendingAccountId;
    }

    var resp = await browser.runtime.sendMessage(msg);
    if (resp.success) {
      showStatus("Account added and syncing!", "success");
      addAccountForm.classList.remove("visible");
      btnShowAddAccount.style.display = "";
      deviceCodeBox.classList.remove("visible");
      // Clear form
      document.getElementById("newEmail").value = "";
      document.getElementById("newUsername").value = "";
      document.getElementById("newPassword").value = "";
      document.getElementById("newCalName").value = "";
      document.getElementById("newEwsUrl").value = "";
      oauthReady = false;
      await refreshAccountList();
    } else {
      showStatus("Failed: " + (resp.error || "Unknown error"), "error");
    }
  } catch (e) {
    showStatus("Error: " + e.message, "error");
  } finally {
    btnAddAccount.disabled = false;
    btnAddAccount.textContent = "Save & Connect";
  }
});

// Render account list
async function refreshAccountList() {
  try {
    var accounts = await browser.runtime.sendMessage({ action: "getAccounts" });
    accountListEl.innerHTML = "";

    if (accounts.length === 0) {
      accountListEl.innerHTML = '<p style="color:var(--text-secondary);font-size:0.9em;margin-bottom:12px">No accounts configured. Click "Add Account" to get started.</p>';
      return;
    }

    for (var acct of accounts) {
      var div = document.createElement("div");
      div.className = "account-card";

      var statusClass = acct.connected ? "connected" : (acct.lastError ? "error" : "disconnected");
      var statusText = acct.connected ? "Connected" : (acct.lastError ? "Error" : "Disconnected");
      var details = acct.calendarName || "Exchange Calendar";
      if (acct.status && acct.status.mappedItems) {
        details += " \u2022 " + acct.status.mappedItems + " items synced";
      }
      if (acct.lastError) {
        details += " \u2022 " + acct.lastError;
      }

      div.innerHTML =
        '<div style="display:flex;justify-content:space-between;align-items:center">' +
          '<div>' +
            '<span class="email">' + escapeHtml(acct.email) + '</span> ' +
            '<span class="status-badge ' + statusClass + '">' + statusText + '</span>' +
            '<div class="details">' + escapeHtml(details) + '</div>' +
          '</div>' +
        '</div>' +
        '<div class="btn-group">' +
          '<button class="btn-secondary btn-sm btn-sync" data-id="' + acct.id + '">Sync Now</button>' +
          '<button class="btn-secondary btn-sm btn-fullsync" data-id="' + acct.id + '">Full Re-sync</button>' +
          '<button class="btn-danger btn-sm btn-remove" data-id="' + acct.id + '">Remove</button>' +
        '</div>';

      accountListEl.appendChild(div);
    }

    // Attach event listeners
    var syncBtns = document.querySelectorAll(".btn-sync");
    for (var i = 0; i < syncBtns.length; i++) {
      syncBtns[i].addEventListener("click", handleSync);
    }
    var fullSyncBtns = document.querySelectorAll(".btn-fullsync");
    for (var j = 0; j < fullSyncBtns.length; j++) {
      fullSyncBtns[j].addEventListener("click", handleFullSync);
    }
    var removeBtns = document.querySelectorAll(".btn-remove");
    for (var k = 0; k < removeBtns.length; k++) {
      removeBtns[k].addEventListener("click", handleRemove);
    }
  } catch (e) {
    console.error("Failed to load accounts:", e);
  }
}

async function handleSync(e) {
  var id = e.target.getAttribute("data-id");
  e.target.textContent = "Syncing...";
  e.target.disabled = true;
  try {
    var resp = await browser.runtime.sendMessage({ action: "syncAccount", accountId: id });
    if (resp.success) {
      showStatus("Sync complete: " + (resp.result.created || 0) + " new, " + (resp.result.updated || 0) + " updated", "success");
    } else {
      showStatus("Sync failed: " + resp.error, "error");
    }
  } catch (err) {
    showStatus("Sync error: " + err.message, "error");
  }
  await refreshAccountList();
}

async function handleFullSync(e) {
  var id = e.target.getAttribute("data-id");
  if (!confirm("Re-download all items from Exchange for this account?")) return;
  e.target.textContent = "Syncing...";
  e.target.disabled = true;
  try {
    var resp = await browser.runtime.sendMessage({ action: "fullSyncAccount", accountId: id });
    if (resp.success) {
      showStatus("Full sync: " + resp.result.total + " items", "success");
    } else {
      showStatus("Full sync failed: " + resp.error, "error");
    }
  } catch (err) {
    showStatus("Full sync error: " + err.message, "error");
  }
  await refreshAccountList();
}

async function handleRemove(e) {
  var id = e.target.getAttribute("data-id");
  var removeCalendar = confirm("Remove this account?\n\nClick OK to also delete the local calendar.\nClick Cancel to keep local data.");
  try {
    await browser.runtime.sendMessage({ action: "removeAccount", accountId: id, removeCalendar: removeCalendar });
    showStatus("Account removed.", "info");
    await refreshAccountList();
  } catch (err) {
    showStatus("Remove error: " + err.message, "error");
  }
}

function escapeHtml(str) {
  var div = document.createElement("div");
  div.appendChild(document.createTextNode(str || ""));
  return div.innerHTML;
}

// =============================================
// Pending Meeting Invitations
// =============================================

var pendingSection = document.getElementById("pendingInvitesSection");
var pendingListEl = document.getElementById("pendingInvitesList");

async function refreshPendingInvites() {
  try {
    var invites = await browser.runtime.sendMessage({ action: "getPendingInvites" });
    var keys = Object.keys(invites);

    if (keys.length === 0) {
      pendingSection.style.display = "none";
      return;
    }

    pendingSection.style.display = "block";
    pendingListEl.innerHTML = "";

    // Fetch all calendars for per-invite picker
    var allCalendars = [];
    try {
      allCalendars = await browser.runtime.sendMessage({ action: "getAllCalendars" });
    } catch (e) { /* use empty */ }

    for (var notifId of keys) {
      var inv = invites[notifId];
      var p = inv.parsed;
      var card = document.createElement("div");
      card.className = "account-card";

      var timeStr = p.dtstart || "";
      if (p.dtend) timeStr += " \u2013 " + p.dtend;
      var from = p.organizerName || p.organizer || "Unknown";

      // Build calendar <select> options
      var calOptions = '';
      for (var ci = 0; ci < allCalendars.length; ci++) {
        var cal = allCalendars[ci];
        var selected = cal.id === inv.calendarId ? ' selected' : '';
        calOptions += '<option value="' + escapeHtml(cal.id) + '"' + selected + '>' + escapeHtml(cal.name) + '</option>';
      }

      card.innerHTML =
        '<div>' +
          '<span class="email">' + escapeHtml(p.summary || "No Subject") + '</span>' +
          '<div class="details">' +
            escapeHtml("From: " + from) + '<br>' +
            escapeHtml("When: " + timeStr) +
            (p.location ? '<br>' + escapeHtml("Where: " + p.location) : '') +
          '</div>' +
        '</div>' +
        '<div style="margin-top:6px">' +
          '<label style="font-size:0.85em;margin-bottom:2px">Add to calendar:</label>' +
          '<select class="invite-cal-select" data-notif="' + escapeHtml(notifId) + '" style="width:100%;padding:5px 8px;border:1px solid var(--input-border);border-radius:4px;font-size:0.85em;background:var(--input-bg);color:var(--text);margin-bottom:6px">' +
            calOptions +
          '</select>' +
        '</div>' +
        '<div class="btn-group">' +
          '<button class="btn-primary btn-sm btn-accept-invite" data-notif="' + escapeHtml(notifId) + '">Accept</button>' +
          '<button class="btn-secondary btn-sm btn-dismiss-invite" data-notif="' + escapeHtml(notifId) + '">Dismiss</button>' +
        '</div>';

      pendingListEl.appendChild(card);
    }

    // Attach listeners
    var acceptBtns = document.querySelectorAll(".btn-accept-invite");
    for (var i = 0; i < acceptBtns.length; i++) {
      acceptBtns[i].addEventListener("click", async function (e) {
        var nid = e.target.getAttribute("data-notif");
        var calSelect = document.querySelector('.invite-cal-select[data-notif="' + nid + '"]');
        var calId = calSelect ? calSelect.value : undefined;
        e.target.textContent = "Adding...";
        e.target.disabled = true;
        await browser.runtime.sendMessage({ action: "acceptInvite", notifId: nid, calendarId: calId });
        showStatus("Meeting added to calendar.", "success");
        await refreshPendingInvites();
      });
    }
    var dismissBtns = document.querySelectorAll(".btn-dismiss-invite");
    for (var j = 0; j < dismissBtns.length; j++) {
      dismissBtns[j].addEventListener("click", async function (e) {
        var nid = e.target.getAttribute("data-notif");
        await browser.runtime.sendMessage({ action: "dismissInvite", notifId: nid });
        await refreshPendingInvites();
      });
    }
  } catch (e) {
    console.error("Failed to load pending invites:", e);
  }
}

refreshAccountList();
refreshPendingInvites();

// =============================================
// iMIP Per-Account Calendar Mapping
// =============================================

var imipAccountMapList = document.getElementById("imipAccountMapList");
var btnSaveImip = document.getElementById("btnSaveImip");
var imipSaveStatus = document.getElementById("imipSaveStatus");

async function loadImipCalendarSetting() {
  try {
    var mailAccounts = await browser.runtime.sendMessage({ action: "getMailAccounts" });
    var calendars = await browser.runtime.sendMessage({ action: "getAllCalendars" });
    var currentMap = await browser.runtime.sendMessage({ action: "getImipAccountMap" });

    if (!mailAccounts || mailAccounts.length === 0) {
      imipAccountMapList.innerHTML = '<p class="hint">No mail accounts found.</p>';
      return;
    }

    imipAccountMapList.innerHTML = "";

    for (var ma of mailAccounts) {
      var row = document.createElement("div");
      row.style.cssText = "margin-bottom:8px;display:flex;align-items:center;gap:8px;flex-wrap:wrap";

      var label = document.createElement("label");
      label.style.cssText = "min-width:180px;font-size:0.9em;flex-shrink:0";
      label.textContent = ma.email;
      label.title = ma.name || ma.email;

      var select = document.createElement("select");
      select.className = "imip-account-select";
      select.setAttribute("data-account-id", ma.id);
      select.style.cssText = "flex:1;min-width:160px;padding:5px 8px;border:1px solid var(--input-border);border-radius:4px;font-size:0.85em;background:var(--input-bg);color:var(--text)";

      var disabledOpt = document.createElement("option");
      disabledOpt.value = "disabled";
      disabledOpt.textContent = "Disabled (Thunderbird handles it)";
      select.appendChild(disabledOpt);

      for (var c of calendars) {
        var opt = document.createElement("option");
        opt.value = c.id;
        opt.textContent = c.name + (c.managed ? " (Exchange)" : "");
        select.appendChild(opt);
      }

      // Pre-select current mapping
      var currentVal = currentMap[ma.id] || "disabled";
      select.value = currentVal;
      // If the stored value is not in the list (calendar deleted), fall back to disabled
      if (select.value !== currentVal) select.value = "disabled";

      row.appendChild(label);
      row.appendChild(select);
      imipAccountMapList.appendChild(row);
    }
  } catch (e) {
    console.error("Failed to load iMIP account mapping:", e);
    imipAccountMapList.innerHTML = '<p class="hint" style="color:var(--error)">Failed to load mail accounts.</p>';
  }
}

btnSaveImip.addEventListener("click", async function () {
  btnSaveImip.disabled = true;
  btnSaveImip.textContent = "Saving...";
  try {
    var selects = document.querySelectorAll(".imip-account-select");
    var accountMap = {};
    for (var i = 0; i < selects.length; i++) {
      var accountId = selects[i].getAttribute("data-account-id");
      accountMap[accountId] = selects[i].value;
    }
    await browser.runtime.sendMessage({ action: "setImipAccountMap", accountMap: accountMap });
    imipSaveStatus.style.display = "inline";
    setTimeout(function () { imipSaveStatus.style.display = "none"; }, 2000);
  } catch (e) {
    showStatus("Failed to save iMIP mapping: " + e.message, "error");
  }
  btnSaveImip.disabled = false;
  btnSaveImip.textContent = "Save";
});

loadImipCalendarSetting();
