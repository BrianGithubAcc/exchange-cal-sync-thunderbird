const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const root = path.resolve(__dirname, "..");

function loadGlobals(files, names) {
  const context = { console, TextEncoder, URL };
  vm.createContext(context);
  for (const file of files) {
    const source = fs.readFileSync(path.join(root, file), "utf8");
    vm.runInContext(source, context, { filename: file });
  }
  const exports = names.map((name) => `${name}: typeof ${name} !== "undefined" ? ${name} : undefined`);
  vm.runInContext(`this.__exports = { ${exports.join(", ")} };`, context);
  return context.__exports;
}

const xml = loadGlobals(["ews/xml-builder.js"], [
  "xmlEscape",
  "buildCreateItemRequest",
  "buildUpdateItemRequest",
  "buildDeleteItemRequest",
  "buildFindItemRequest",
]);

const sync = loadGlobals(["sync/engine.js"], [
  "ewsItemToIcs",
  "icsToEwsFields",
  "stableEwsUid",
]);

const policy = loadGlobals(["ews/client.js"], [
  "validateEwsUrl",
  "DEFAULT_EWS_URL",
]);

test("XML escaping protects EWS request values", () => {
  assert.equal(xml.xmlEscape("A&B <test> \"quote\""), "A&amp;B &lt;test&gt; &quot;quote&quot;");
});

test("CreateItem contains the calendar event fields", () => {
  const request = xml.buildCreateItemRequest({
    subject: "Planning & review",
    body: "Details",
    start: "2026-10-10T10:00:00Z",
    end: "2026-10-10T11:00:00Z",
    location: "Room 1",
    isAllDay: false,
  });

  assert.match(request, /<m:CreateItem SendMeetingInvitations="SendToNone">/);
  assert.match(request, /Planning &amp; review/);
  assert.match(request, /<t:Start>2026-10-10T10:00:00Z<\/t:Start>/);
});

test("UpdateItem carries the EWS item identity and changed fields", () => {
  const request = xml.buildUpdateItemRequest("item-id", "change-key", {
    subject: "Updated",
    location: "Room 2",
  });

  assert.match(request, /Id="item-id" ChangeKey="change-key"/);
  assert.match(request, /<t:Subject>Updated<\/t:Subject>/);
  assert.match(request, /<t:Location>Room 2<\/t:Location>/);
});

test("DeleteItem moves the event to Deleted Items", () => {
  const request = xml.buildDeleteItemRequest("item-id", "change-key");
  assert.match(request, /DeleteType="MoveToDeletedItems"/);
  assert.match(request, /SendMeetingCancellations="SendToNone"/);
});

test("EWS calendar results map to iCalendar", () => {
  const ics = sync.ewsItemToIcs({
    id: "ews-id",
    subject: "A meeting",
    body: "Description",
    start: "2026-10-10T10:00:00Z",
    end: "2026-10-10T11:00:00Z",
    location: "Room 1",
    isAllDay: false,
    reminderSet: true,
    reminderMinutes: 15,
  });

  assert.match(ics, /BEGIN:VCALENDAR/);
  assert.match(ics, /SUMMARY:A meeting/);
  assert.match(ics, /DTSTART:20261010T100000Z/);
  assert.match(ics, /TRIGGER:-PT15M/);
});

test("iCalendar changes map back to EWS fields", () => {
  const fields = sync.icsToEwsFields([
    "BEGIN:VCALENDAR",
    "BEGIN:VEVENT",
    "SUMMARY:Edited event",
    "DESCRIPTION:Updated description",
    "LOCATION:Room 3",
    "DTSTART:20261010T120000Z",
    "DTEND:20261010T130000Z",
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n"));

  assert.deepEqual(JSON.parse(JSON.stringify(fields)), {
    subject: "Edited event",
    body: "Updated description",
    location: "Room 3",
    start: "2026-10-10T12:00:00Z",
    end: "2026-10-10T13:00:00Z",
    isAllDay: false,
  });
});

test("all-day iCalendar events retain date-only values", () => {
  const fields = sync.icsToEwsFields([
    "BEGIN:VEVENT",
    "DTSTART;VALUE=DATE:20261010",
    "DTEND;VALUE=DATE:20261011",
    "END:VEVENT",
  ].join("\r\n"));

  assert.equal(fields.start, "2026-10-10T00:00:00Z");
  assert.equal(fields.end, "2026-10-11T00:00:00Z");
  assert.equal(fields.isAllDay, true);
});

test("EWS UIDs are injective and organizer input cannot inject ICS lines", () => {
  assert.notEqual(sync.stableEwsUid("Aa"), sync.stableEwsUid("BB"));
  const ics = sync.ewsItemToIcs({
    id: "item-id",
    subject: "Safe",
    start: "2026-10-10T10:00:00Z",
    end: "2026-10-10T11:00:00Z",
    organizer: { name: "Organizer", email: "attacker@example.com\r\nX-Injected: yes" },
  });
  assert.doesNotMatch(ics, /X-Injected/);
});

test("EWS endpoint policy rejects credential-forwarding destinations", () => {
  assert.equal(policy.validateEwsUrl(""), policy.DEFAULT_EWS_URL);
  assert.equal(policy.validateEwsUrl(policy.DEFAULT_EWS_URL), policy.DEFAULT_EWS_URL);
  assert.throws(() => policy.validateEwsUrl("https://example.com/EWS/Exchange.asmx"));
  assert.throws(() => policy.validateEwsUrl("https://outlook.office365.com/EWS/Exchange.asmx?next=https://example.com"));
  assert.throws(() => policy.validateEwsUrl("http://outlook.office365.com/EWS/Exchange.asmx"));
});
