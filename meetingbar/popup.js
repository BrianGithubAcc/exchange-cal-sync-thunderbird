"use strict";

var loadingEl = document.getElementById("loading");
var notMeetingEl = document.getElementById("notMeeting");
var contentEl = document.getElementById("content");
var subjectEl = document.getElementById("subject");
var detailsEl = document.getElementById("details");
var statusEl = document.getElementById("status");
var btnAccept = document.getElementById("btnAccept");
var btnTentative = document.getElementById("btnTentative");
var btnDecline = document.getElementById("btnDecline");

var currentMeeting = null;

function showStatus(msg, type) {
  statusEl.textContent = msg;
  statusEl.className = "visible " + (type || "info");
}

async function init() {
  try {
    // Get the currently displayed message
    var tabs = await browser.tabs.query({ active: true, currentWindow: true });
    if (!tabs.length) { showNotMeeting(); return; }

    var msg = await browser.messageDisplay.getDisplayedMessage(tabs[0].id);
    if (!msg) { showNotMeeting(); return; }

    // Ask background to check if this is a meeting and get details
    var result = await browser.runtime.sendMessage({
      action: "checkMeetingMessage",
      messageId: msg.id,
      subject: msg.subject,
      author: msg.author,
    });

    if (!result || !result.isMeeting) {
      showNotMeeting();
      return;
    }

    currentMeeting = result;
    subjectEl.textContent = result.subject || msg.subject || "Meeting";
    var info = "";
    if (result.organizer) info += "From: " + result.organizer + "\n";
    if (result.start) info += "When: " + result.start;
    if (result.end) info += " - " + result.end;
    if (result.location) info += "\nWhere: " + result.location;
    detailsEl.textContent = info;

    loadingEl.style.display = "none";
    contentEl.style.display = "block";
  } catch (e) {
    loadingEl.textContent = "Error: " + e.message;
  }
}

function showNotMeeting() {
  loadingEl.style.display = "none";
  notMeetingEl.style.display = "block";
}

async function respond(responseType) {
  if (!currentMeeting) return;

  btnAccept.disabled = true;
  btnTentative.disabled = true;
  btnDecline.disabled = true;
  showStatus("Sending response...", "info");

  try {
    var result = await browser.runtime.sendMessage({
      action: "respondToMeeting",
      responseType: responseType,
      itemId: currentMeeting.itemId,
      changeKey: currentMeeting.changeKey,
      accountId: currentMeeting.accountId,
      messageId: currentMeeting.messageId,
    });

    if (result.success) {
      var labels = { accept: "Accepted", tentative: "Tentatively accepted", decline: "Declined" };
      showStatus(labels[responseType] + "!", "success");
      // Close popup after a moment
      setTimeout(function () { window.close(); }, 1500);
    } else {
      showStatus("Failed: " + (result.error || "Unknown error"), "error");
      btnAccept.disabled = false;
      btnTentative.disabled = false;
      btnDecline.disabled = false;
    }
  } catch (e) {
    showStatus("Error: " + e.message, "error");
    btnAccept.disabled = false;
    btnTentative.disabled = false;
    btnDecline.disabled = false;
  }
}

btnAccept.addEventListener("click", function () { respond("accept"); });
btnTentative.addEventListener("click", function () { respond("tentative"); });
btnDecline.addEventListener("click", function () { respond("decline"); });

init();
