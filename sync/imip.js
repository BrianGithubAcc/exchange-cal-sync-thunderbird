console.log("[ExCalSync] imip.js loading");
/**
 * iMIP (iCalendar Message-Based Interoperability Protocol) parser.
 * Parses text/calendar email parts to extract meeting request/cancellation data.
 * Pure string parsing of RFC 5545 iCalendar format — no external dependencies.
 */

/**
 * Check if a content type indicates an iCalendar invite.
 */
function isCalendarInvite(contentType) {
  if (!contentType) return false;
  var ct = contentType.toLowerCase();
  return ct.indexOf("text/calendar") !== -1 || ct.indexOf("application/ics") !== -1;
}

/**
 * Parse iCalendar text and extract meeting-relevant fields.
 * Returns an object with method, uid, summary, dtstart, dtend, organizer, etc.
 * Returns null if the text is not valid iCalendar.
 */
function parseICalendar(icsText) {
  if (!icsText || icsText.indexOf("BEGIN:VCALENDAR") === -1) return null;

  // Unfold lines (RFC 5545 §3.1: continuation lines start with space or tab)
  var unfolded = icsText.replace(/\r\n[ \t]/g, "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  var lines = unfolded.split("\n");

  var result = {
    method: null,
    uid: null,
    summary: null,
    description: null,
    location: null,
    dtstart: null,
    dtend: null,
    organizer: null,
    organizerName: null,
    attendees: [],
    sequence: 0,
    status: null,
    rawIcs: icsText,
  };

  var inVevent = false;

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (!line) continue;

    var colonIdx = line.indexOf(":");
    if (colonIdx === -1) continue;

    var keyPart = line.substring(0, colonIdx);
    var value = line.substring(colonIdx + 1);
    var key = keyPart.split(";")[0].toUpperCase();

    // Track VEVENT boundaries
    if (key === "BEGIN" && value.toUpperCase() === "VEVENT") { inVevent = true; continue; }
    if (key === "END" && value.toUpperCase() === "VEVENT") { inVevent = false; continue; }

    // Calendar-level properties
    if (!inVevent) {
      if (key === "METHOD") {
        result.method = value.trim().toUpperCase();
      }
      continue;
    }

    // VEVENT-level properties
    switch (key) {
      case "UID":
        result.uid = value.trim();
        break;
      case "SUMMARY":
        result.summary = imipUnescapeValue(value);
        break;
      case "DESCRIPTION":
        result.description = imipUnescapeValue(value);
        break;
      case "LOCATION":
        result.location = imipUnescapeValue(value);
        break;
      case "DTSTART":
        result.dtstart = imipDateToReadable(value.trim());
        break;
      case "DTEND":
        result.dtend = imipDateToReadable(value.trim());
        break;
      case "ORGANIZER":
        result.organizer = imipExtractEmail(value);
        result.organizerName = imipExtractParam(keyPart, "CN");
        break;
      case "ATTENDEE":
        result.attendees.push({
          email: imipExtractEmail(value),
          name: imipExtractParam(keyPart, "CN"),
          role: imipExtractParam(keyPart, "ROLE"),
          partstat: imipExtractParam(keyPart, "PARTSTAT"),
        });
        break;
      case "SEQUENCE":
        result.sequence = parseInt(value.trim(), 10) || 0;
        break;
      case "STATUS":
        result.status = value.trim().toUpperCase();
        break;
    }
  }

  // Must have at least a UID to be useful
  if (!result.uid) return null;

  return result;
}

/**
 * Unescape iCalendar text values (RFC 5545 §3.3.11).
 */
function imipUnescapeValue(str) {
  if (!str) return "";
  return str
    .replace(/\\n/gi, "\n")
    .replace(/\\,/g, ",")
    .replace(/\\;/g, ";")
    .replace(/\\\\/g, "\\");
}

/**
 * Extract email address from an iCalendar value like "mailto:user@example.com".
 */
function imipExtractEmail(value) {
  if (!value) return null;
  var lower = value.toLowerCase().trim();
  if (lower.indexOf("mailto:") === 0) {
    return value.trim().substring(7).toLowerCase();
  }
  // Might just be a bare email
  if (value.indexOf("@") !== -1) {
    return value.trim().toLowerCase();
  }
  return null;
}

/**
 * Extract a parameter value from an iCalendar property key part.
 * e.g. from "ATTENDEE;CN=John Doe;ROLE=REQ-PARTICIPANT", extract "CN" -> "John Doe"
 */
function imipExtractParam(keyPart, paramName) {
  if (!keyPart) return null;
  var parts = keyPart.split(";");
  for (var i = 1; i < parts.length; i++) {
    var eqIdx = parts[i].indexOf("=");
    if (eqIdx === -1) continue;
    var pName = parts[i].substring(0, eqIdx).toUpperCase();
    if (pName === paramName.toUpperCase()) {
      var pVal = parts[i].substring(eqIdx + 1);
      // Remove surrounding quotes
      if (pVal.length >= 2 && pVal.charAt(0) === '"' && pVal.charAt(pVal.length - 1) === '"') {
        pVal = pVal.substring(1, pVal.length - 1);
      }
      return pVal;
    }
  }
  return null;
}

/**
 * Convert an iCalendar date string to a human-readable format.
 * Input: "20260515T140000Z" or "20260515" (all-day)
 * Output: "2026-05-15 14:00 UTC" or "2026-05-15"
 */
function imipDateToReadable(icsDate) {
  if (!icsDate) return null;
  // Strip trailing Z for parsing
  var clean = icsDate.replace("Z", "");
  if (clean.length === 8) {
    // All-day: YYYYMMDD
    return clean.substring(0, 4) + "-" + clean.substring(4, 6) + "-" + clean.substring(6, 8);
  }
  if (clean.length >= 15) {
    // DateTime: YYYYMMDDTHHMMSS
    var dateStr = clean.substring(0, 4) + "-" + clean.substring(4, 6) + "-" + clean.substring(6, 8);
    var timeStr = clean.substring(9, 11) + ":" + clean.substring(11, 13);
    var suffix = icsDate.indexOf("Z") !== -1 ? " UTC" : "";
    return dateStr + " " + timeStr + suffix;
  }
  // Unknown format, return as-is
  return icsDate;
}

/**
 * Extract the domain from an email address.
 */
function imipGetDomain(email) {
  if (!email) return null;
  var atIdx = email.indexOf("@");
  if (atIdx === -1) return null;
  return email.substring(atIdx + 1).toLowerCase();
}

// Globals: isCalendarInvite, parseICalendar, imipGetDomain
