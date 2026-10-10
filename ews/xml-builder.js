console.log("[ExCalSync] xml-builder.js loading");
/**
 * XML builder utilities for EWS SOAP requests.
 * Constructs well-formed XML strings without innerHTML or DOM parsing.
 */

const EWS_NAMESPACES = {
  soap: "http://schemas.xmlsoap.org/soap/envelope/",
  t: "http://schemas.microsoft.com/exchange/services/2006/types",
  m: "http://schemas.microsoft.com/exchange/services/2006/messages",
};

/**
 * Escapes a string for safe inclusion in XML.
 */
function xmlEscape(str) {
  if (str == null) return "";
  return String(str)
    // Strip control characters illegal in XML 1.0 (keep \t, \n, \r)
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Wraps body XML in a SOAP envelope with EWS namespaces.
 */
function soapEnvelope(bodyXml, headerXml = "") {
  return [
    '<?xml version="1.0" encoding="utf-8"?>',
    `<soap:Envelope xmlns:soap="${EWS_NAMESPACES.soap}"`,
    `  xmlns:t="${EWS_NAMESPACES.t}"`,
    `  xmlns:m="${EWS_NAMESPACES.m}">`,
    headerXml ? `<soap:Header>${headerXml}</soap:Header>` : "",
    `<soap:Body>${bodyXml}</soap:Body>`,
    "</soap:Envelope>",
  ].join("\n");
}

function buildGetCalendarFolderRequest() {
  return soapEnvelope(`
    <m:GetFolder>
      <m:FolderShape><t:BaseShape>Default</t:BaseShape></m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="calendar"/>
      </m:FolderIds>
    </m:GetFolder>
  `);
}

/**
 * Builds FindItem request to list calendar items.
 */
function buildFindItemRequest(startDate, endDate) {
  const startIso = startDate.toISOString();
  const endIso = endDate.toISOString();

  return soapEnvelope(`
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="item:DateTimeReceived"/>
          <t:FieldURI FieldURI="item:LastModifiedTime"/>
          <t:FieldURI FieldURI="calendar:Start"/>
          <t:FieldURI FieldURI="calendar:End"/>
          <t:FieldURI FieldURI="calendar:Location"/>
          <t:FieldURI FieldURI="item:Body"/>
          <t:FieldURI FieldURI="calendar:IsAllDayEvent"/>
          <t:FieldURI FieldURI="calendar:Organizer"/>
          <t:FieldURI FieldURI="calendar:RequiredAttendees"/>
          <t:FieldURI FieldURI="calendar:OptionalAttendees"/>
          <t:FieldURI FieldURI="item:Sensitivity"/>
          <t:FieldURI FieldURI="item:Importance"/>
          <t:FieldURI FieldURI="calendar:LegacyFreeBusyStatus"/>
          <t:FieldURI FieldURI="item:Categories"/>
          <t:FieldURI FieldURI="item:ReminderMinutesBeforeStart"/>
          <t:FieldURI FieldURI="item:ReminderIsSet"/>
          <t:FieldURI FieldURI="calendar:Recurrence"/>
          <t:FieldURI FieldURI="item:HasAttachments"/>
          <t:FieldURI FieldURI="calendar:IsCancelled"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:CalendarView StartDate="${xmlEscape(startIso)}" EndDate="${xmlEscape(endIso)}" MaxEntriesReturned="500"/>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="calendar"/>
      </m:ParentFolderIds>
    </m:FindItem>
  `);
}

/**
 * Builds GetItem request to get full details of specific items.
 */
function buildGetItemRequest(itemIds) {
  const idElements = itemIds
    .map((id) => `<t:ItemId Id="${xmlEscape(id.id)}"${id.changeKey ? ` ChangeKey="${xmlEscape(id.changeKey)}"` : ""}/>`)
    .join("\n");

  return soapEnvelope(`
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>AllProperties</t:BaseShape>
        <t:BodyType>HTML</t:BodyType>
      </m:ItemShape>
      <m:ItemIds>
        ${idElements}
      </m:ItemIds>
    </m:GetItem>
  `);
}

/**
 * Builds SyncFolderItems request for incremental sync.
 */
function buildSyncFolderItemsRequest(syncState = null) {
  const syncStateEl = syncState
    ? `<m:SyncState>${xmlEscape(syncState)}</m:SyncState>`
    : "";

  return soapEnvelope(`
    <m:SyncFolderItems>
      <m:ItemShape>
        <t:BaseShape>AllProperties</t:BaseShape>
      </m:ItemShape>
      <m:SyncFolderId>
        <t:DistinguishedFolderId Id="calendar"/>
      </m:SyncFolderId>
      ${syncStateEl}
      <m:MaxChangesReturned>512</m:MaxChangesReturned>
    </m:SyncFolderItems>
  `);
}

/**
 * Builds CreateItem request for a new calendar event.
 */
function buildCreateItemRequest(event) {
  const attendees = (event.attendees || [])
    .map(
      (a) => `
      <t:Attendee>
        <t:Mailbox>
          <t:EmailAddress>${xmlEscape(a.email)}</t:EmailAddress>
          <t:Name>${xmlEscape(a.name || a.email)}</t:Name>
        </t:Mailbox>
      </t:Attendee>`
    )
    .join("");

  return soapEnvelope(`
    <m:CreateItem SendMeetingInvitations="SendToNone">
      <m:Items>
        <t:CalendarItem>
          <t:Subject>${xmlEscape(event.subject)}</t:Subject>
          <t:Body BodyType="Text">${xmlEscape(event.body || "")}</t:Body>
          <t:Start>${xmlEscape(event.start)}</t:Start>
          <t:End>${xmlEscape(event.end)}</t:End>
          <t:Location>${xmlEscape(event.location || "")}</t:Location>
          <t:IsAllDayEvent>${event.isAllDay ? "true" : "false"}</t:IsAllDayEvent>
          <t:ReminderIsSet>${event.reminderSet ? "true" : "false"}</t:ReminderIsSet>
          ${event.reminderSet && event.reminderMinutes != null
            ? `<t:ReminderMinutesBeforeStart>${xmlEscape(event.reminderMinutes)}</t:ReminderMinutesBeforeStart>`
            : ""}
          ${event.sensitivity && event.sensitivity !== "Normal"
            ? `<t:Sensitivity>${xmlEscape(event.sensitivity)}</t:Sensitivity>`
            : ""}
          ${
            attendees
              ? `<t:RequiredAttendees>${attendees}</t:RequiredAttendees>`
              : ""
          }
        </t:CalendarItem>
      </m:Items>
    </m:CreateItem>
  `);
}

/**
 * Builds UpdateItem request for modifying an existing calendar event.
 */
function buildUpdateItemRequest(itemId, changeKey, updates) {
  const setFields = [];

  if (updates.subject !== undefined) {
    setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="item:Subject"/>
        <t:CalendarItem><t:Subject>${xmlEscape(updates.subject)}</t:Subject></t:CalendarItem>
      </t:SetItemField>
    `);
  }
  if (updates.body !== undefined) {
    setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="item:Body"/>
        <t:CalendarItem><t:Body BodyType="Text">${xmlEscape(updates.body)}</t:Body></t:CalendarItem>
      </t:SetItemField>
    `);
  }
  if (updates.start !== undefined) {
    setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="calendar:Start"/>
        <t:CalendarItem><t:Start>${xmlEscape(updates.start)}</t:Start></t:CalendarItem>
      </t:SetItemField>
    `);
  }
  if (updates.end !== undefined) {
    setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="calendar:End"/>
        <t:CalendarItem><t:End>${xmlEscape(updates.end)}</t:End></t:CalendarItem>
      </t:SetItemField>
    `);
  }
  if (updates.location !== undefined) {
    setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="calendar:Location"/>
        <t:CalendarItem><t:Location>${xmlEscape(updates.location)}</t:Location></t:CalendarItem>
      </t:SetItemField>
    `);
  }
  if (updates.isAllDay !== undefined) {
    setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="calendar:IsAllDayEvent"/>
        <t:CalendarItem><t:IsAllDayEvent>${updates.isAllDay ? "true" : "false"}</t:IsAllDayEvent></t:CalendarItem>
      </t:SetItemField>
    `);
  }
  if (updates.reminderSet !== undefined) {
    setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="item:ReminderIsSet"/>
        <t:CalendarItem><t:ReminderIsSet>${updates.reminderSet ? "true" : "false"}</t:ReminderIsSet></t:CalendarItem>
      </t:SetItemField>
    `);
  }
  if (updates.reminderMinutes !== undefined) {
    setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="item:ReminderMinutesBeforeStart"/>
        <t:CalendarItem><t:ReminderMinutesBeforeStart>${xmlEscape(updates.reminderMinutes)}</t:ReminderMinutesBeforeStart></t:CalendarItem>
      </t:SetItemField>
    `);
  }
  if (updates.sensitivity !== undefined) {
    setFields.push(`
      <t:SetItemField>
        <t:FieldURI FieldURI="item:Sensitivity"/>
        <t:CalendarItem><t:Sensitivity>${xmlEscape(updates.sensitivity)}</t:Sensitivity></t:CalendarItem>
      </t:SetItemField>
    `);
  }

  return soapEnvelope(`
    <m:UpdateItem ConflictResolution="NeverOverwrite" SendMeetingInvitationsOrCancellations="SendToNone">
      <m:ItemChanges>
        <t:ItemChange>
          <t:ItemId Id="${xmlEscape(itemId)}" ChangeKey="${xmlEscape(changeKey)}"/>
          <t:Updates>
            ${setFields.join("")}
          </t:Updates>
        </t:ItemChange>
      </m:ItemChanges>
    </m:UpdateItem>
  `);
}

/**
 * Builds DeleteItem request.
 */
function buildDeleteItemRequest(itemId, changeKey) {
  return soapEnvelope(`
    <m:DeleteItem DeleteType="MoveToDeletedItems" SendMeetingCancellations="SendToNone">
      <m:ItemIds>
        <t:ItemId Id="${xmlEscape(itemId)}" ChangeKey="${xmlEscape(changeKey)}"/>
      </m:ItemIds>
    </m:DeleteItem>
  `);
}

/**
 * Builds a meeting response request (Accept/Tentative/Decline).
 * @param {string} itemId - The CalendarItem ItemId
 * @param {string} changeKey - The CalendarItem ChangeKey
 * @param {string} responseType - "accept", "tentative", or "decline"
 */
function buildMeetingResponseRequest(itemId, changeKey, responseType) {
  var elementMap = {
    accept: "AcceptItem",
    tentative: "TentativelyAcceptItem",
    decline: "DeclineItem",
  };
  var element = elementMap[responseType];
  if (!element) throw new Error("Invalid responseType: " + responseType);

  return soapEnvelope(`
    <m:CreateItem MessageDisposition="SendAndSaveCopy">
      <m:Items>
        <t:${element}>
          <t:ReferenceItemId Id="${xmlEscape(itemId)}" ChangeKey="${xmlEscape(changeKey)}"/>
        </t:${element}>
      </m:Items>
    </m:CreateItem>
  `);
}

/**
 * Builds a FindItem request to search calendar by subject.
 */
function buildFindItemBySubjectRequest(subject, startDate, endDate) {
  var startIso = startDate.toISOString();
  var endIso = endDate.toISOString();

  return soapEnvelope(`
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="item:Subject"/>
          <t:FieldURI FieldURI="calendar:Start"/>
          <t:FieldURI FieldURI="calendar:End"/>
          <t:FieldURI FieldURI="calendar:Location"/>
          <t:FieldURI FieldURI="calendar:Organizer"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:Restriction>
        <t:And>
          <t:Contains ContainmentMode="Substring" ContainmentComparison="IgnoreCase">
            <t:FieldURI FieldURI="item:Subject"/>
            <t:Constant Value="${xmlEscape(subject)}"/>
          </t:Contains>
          <t:IsGreaterThanOrEqualTo>
            <t:FieldURI FieldURI="calendar:Start"/>
            <t:FieldURIOrConstant>
              <t:Constant Value="${xmlEscape(startIso)}"/>
            </t:FieldURIOrConstant>
          </t:IsGreaterThanOrEqualTo>
          <t:IsLessThanOrEqualTo>
            <t:FieldURI FieldURI="calendar:End"/>
            <t:FieldURIOrConstant>
              <t:Constant Value="${xmlEscape(endIso)}"/>
            </t:FieldURIOrConstant>
          </t:IsLessThanOrEqualTo>
        </t:And>
      </m:Restriction>
      <m:IndexedPageItemView MaxEntriesReturned="10" Offset="0" BasePoint="Beginning"/>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="calendar"/>
      </m:ParentFolderIds>
    </m:FindItem>
  `);
}

// Globals: EWS_NAMESPACES, xmlEscape, soapEnvelope, buildAutodiscoverRequest,
// buildFindItemRequest, buildGetItemRequest, buildSyncFolderItemsRequest,
// buildCreateItemRequest, buildUpdateItemRequest, buildDeleteItemRequest,
// buildMeetingResponseRequest, buildFindItemBySubjectRequest
