/* global ExtensionCommon, ChromeUtils, Components, Services */
/* exported CalendarAPI */

"use strict";

var cal;
try {
  ({ cal } = ChromeUtils.importESModule("resource:///modules/calendar/calUtils.sys.mjs"));
  console.log("[ExCalSync API] calUtils loaded successfully");
} catch (e) {
  console.error("[ExCalSync API] Failed to import calUtils:", e);
}

var CalendarAPI = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    const calMgr = cal.manager;
    const MANAGED_PROP = "exchange-cal-sync.managed";
    const ExtensionError = ExtensionCommon.ExtensionError || Error;

    function findCalendar(calendarId) {
      for (const calendar of calMgr.getCalendars()) {
        if (calendar.id === calendarId) {
          return calendar;
        }
      }
      return null;
    }

    function managedCalendar(calendarId) {
      const calendar = findCalendar(calendarId);
      if (!calendar || calendar.getProperty(MANAGED_PROP) !== "true") {
        throw new ExtensionError("Calendar is not managed by Exchange Calendar Sync");
      }
      return calendar;
    }

    function icsEscape(str) {
      if (!str) return "";
      return str.replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
    }

    function itemToIcs(item) {
      try {
        // Prefer Thunderbird's complete serialization. The fallback below is
        // only for older calendar implementations without icalString.
        try {
          if (typeof item.icalString === "string" &&
              item.icalString.length <= 1024 * 1024 &&
              item.icalString.indexOf("BEGIN:VCALENDAR") !== -1) {
            return item.icalString;
          }
        } catch (serializationError) { /* use the compatibility serializer */ }

        var start = item.startDate || item.entryDate;
        var end = item.endDate || item.dueDate;
        var lines = [
          "BEGIN:VCALENDAR",
          "VERSION:2.0",
          "PRODID:-//Exchange Calendar Sync//EN",
          "BEGIN:VEVENT",
        ];

        lines.push("UID:" + (item.id || "unknown"));
        if (item.title) lines.push("SUMMARY:" + icsEscape(item.title));
        if (item.getProperty("DESCRIPTION")) {
          lines.push("DESCRIPTION:" + icsEscape(item.getProperty("DESCRIPTION")));
        }
        if (item.getProperty("LOCATION")) {
          lines.push("LOCATION:" + icsEscape(item.getProperty("LOCATION")));
        }
        if (start) {
          try {
            if (start.isDate) {
              // All-day event: use date-only format (no timezone conversion)
              if (start.icalString) {
                lines.push("DTSTART;VALUE=DATE:" + start.icalString);
              } else if (start.jsDate) {
                lines.push("DTSTART;VALUE=DATE:" + start.jsDate.toISOString().replace(/[-:]/g, "").substring(0, 8));
              }
            } else if (start.jsDate) {
              // Timed event: always use UTC via jsDate to preserve timezone accuracy
              lines.push("DTSTART:" + start.jsDate.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}/, ""));
            } else if (start.icalString) {
              lines.push("DTSTART:" + start.icalString);
            }
          } catch (de) {
            console.warn("[ExCalSync API] itemToIcs: could not serialize start date:", de);
          }
        }
        if (end) {
          try {
            if (end.isDate) {
              // All-day event: use date-only format (no timezone conversion)
              if (end.icalString) {
                lines.push("DTEND;VALUE=DATE:" + end.icalString);
              } else if (end.jsDate) {
                lines.push("DTEND;VALUE=DATE:" + end.jsDate.toISOString().replace(/[-:]/g, "").substring(0, 8));
              }
            } else if (end.jsDate) {
              // Timed event: always use UTC via jsDate to preserve timezone accuracy
              lines.push("DTEND:" + end.jsDate.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}/, ""));
            } else if (end.icalString) {
              lines.push("DTEND:" + end.icalString);
            }
          } catch (de2) {
            console.warn("[ExCalSync API] itemToIcs: could not serialize end date:", de2);
          }
        }

        lines.push("END:VEVENT");
        lines.push("END:VCALENDAR");
        return lines.join("\r\n");
      } catch (e) {
        console.error("[ExCalSync API] itemToIcs error:", e);
        return null;
      }
    }

    function icsToItem(icsData) {
      try {
        var parser = Components.classes["@mozilla.org/calendar/ics-parser;1"]
                               .createInstance(Components.interfaces.calIIcsParser);
        parser.parseString(icsData);
        var items = parser.getItems();
        return items.length > 0 ? items[0] : null;
      } catch (e) {
        console.error("[ExCalSync API] icsToItem error:", e);
        return null;
      }
    }

    // Observer for tracking local changes
    const changeObserver = {
      listeners: new Set(),
      registered: false,

      calendarObserver: {
        QueryInterface: ChromeUtils.generateQI(["calIObserver"]),
        onStartBatch() {},
        onEndBatch() {},
        onLoad() {},
        onError() {},
        onPropertyChanged() {},
        onPropertyDeleting() {},

        onAddItem(item) {
          if (!item.calendar || item.calendar.getProperty(MANAGED_PROP) !== "true") return;
          changeObserver.notify({
            calendarId: String(item.calendar.id),
            itemId: String(item.id),
            changeType: "added",
            icsData: itemToIcs(item) || "",
          });
        },

        onModifyItem(newItem) {
          if (!newItem.calendar || newItem.calendar.getProperty(MANAGED_PROP) !== "true") return;
          changeObserver.notify({
            calendarId: String(newItem.calendar.id),
            itemId: String(newItem.id),
            changeType: "modified",
            icsData: itemToIcs(newItem) || "",
          });
        },

        onDeleteItem(item) {
          if (!item.calendar || item.calendar.getProperty(MANAGED_PROP) !== "true") return;
          changeObserver.notify({
            calendarId: String(item.calendar.id),
            itemId: String(item.id),
            changeType: "deleted",
          });
        },
      },

      register() {
        if (!this.registered) {
          calMgr.addCalendarObserver(this.calendarObserver);
          this.registered = true;
        }
      },

      unregister() {
        if (this.registered && this.listeners.size === 0) {
          calMgr.removeCalendarObserver(this.calendarObserver);
          this.registered = false;
        }
      },

      notify(changeInfo) {
        for (const fire of this.listeners) {
          fire.async(changeInfo);
        }
      },
    };

    context.callOnClose({
      close() {
        if (changeObserver.registered) {
          calMgr.removeCalendarObserver(changeObserver.calendarObserver);
          changeObserver.registered = false;
          changeObserver.listeners.clear();
        }
      },
    });

    return {
      CalendarAPI: {
        async getCalendars() {
          console.log("[ExCalSync API] getCalendars()");
          try {
            var calendars = calMgr.getCalendars();
            var result = [];
            for (var c of calendars) {
              if (c.getProperty(MANAGED_PROP) !== "true") continue;
              result.push({
                id: String(c.id),
                name: String(c.name || ""),
                type: String(c.type || ""),
                color: String(c.getProperty("color") || "#0078D4"),
                readOnly: !!c.readOnly,
                managed: c.getProperty(MANAGED_PROP) === "true",
              });
            }
            console.log("[ExCalSync API] getCalendars: found", result.length);
            return result;
          } catch (e) {
            console.error("[ExCalSync API] getCalendars error:", e);
            throw new ExtensionError("getCalendars failed: " + e.message);
          }
        },

        async createCalendar(name, color) {
          console.log("[ExCalSync API] createCalendar:", name);
          try {
            var calendar = calMgr.createCalendar(
              "storage",
              Services.io.newURI("moz-storage-calendar://")
            );
            calendar.name = name;
            calendar.setProperty("color", color);
            calendar.setProperty(MANAGED_PROP, "true");
            calMgr.registerCalendar(calendar);
            console.log("[ExCalSync API] Created calendar:", calendar.id);
            return {
              id: String(calendar.id),
              name: String(calendar.name),
              type: String(calendar.type),
              color: String(color),
            };
          } catch (e) {
            console.error("[ExCalSync API] createCalendar error:", e);
            throw new ExtensionError("createCalendar failed: " + e.message);
          }
        },

        async removeCalendar(calendarId) {
          console.log("[ExCalSync API] removeCalendar:", calendarId);
          var calendar = managedCalendar(calendarId);
          calMgr.removeCalendar(calendar);
          return true;
        },

        async getItems(calendarId, itemType, rangeStartIso, rangeEndIso) {
          console.log("[ExCalSync API] getItems:", calendarId, itemType);
          var calendar = managedCalendar(calendarId);

          try {
            var filter = Components.interfaces.calICalendar.ITEM_FILTER_TYPE_EVENT |
                         Components.interfaces.calICalendar.ITEM_FILTER_COMPLETED_ALL;

            var rangeStart = null;
            var rangeEnd = null;
            if (rangeStartIso) {
              rangeStart = cal.createDateTime(rangeStartIso.replace(/[-:]/g, "").replace(/\.\d{3}/, ""));
            }
            if (rangeEndIso) {
              rangeEnd = cal.createDateTime(rangeEndIso.replace(/[-:]/g, "").replace(/\.\d{3}/, ""));
            }

            // TB 128+ uses promise-based getItemsAsArray
            var items;
            if (typeof calendar.getItemsAsArray === "function") {
              items = await calendar.getItemsAsArray(filter, 0, rangeStart, rangeEnd);
            } else {
              // Fallback to callback-based API
              items = await new Promise(function (resolve, reject) {
                var results = [];
                calendar.getItems(filter, 0, rangeStart, rangeEnd, {
                  QueryInterface: ChromeUtils.generateQI(["calIOperationListener"]),
                  onGetResult(aCal, aStatus, aItemType, aDetail, aItems) {
                    for (var item of aItems) {
                      results.push(item);
                    }
                  },
                  onOperationComplete(aCal, aStatus) {
                    if (Components.isSuccessCode(aStatus)) {
                      resolve(results);
                    } else {
                      reject(new Error("getItems failed: " + aStatus));
                    }
                  },
                });
              });
            }

            var result = [];
            for (var item of items) {
              var ics = itemToIcs(item);
              var lastMod = null;
              try {
                if (item.lastModifiedTime && item.lastModifiedTime.jsDate) {
                  lastMod = item.lastModifiedTime.jsDate.toISOString();
                }
              } catch (e) { /* some items have invalid lastModifiedTime */ }
              result.push({
                id: String(item.id),
                title: String(item.title || ""),
                icsData: ics || "",
                lastModified: lastMod,
              });
            }
            console.log("[ExCalSync API] getItems: returning", result.length, "items");
            return result;
          } catch (e) {
            console.error("[ExCalSync API] getItems error:", e);
            throw new ExtensionError("getItems failed: " + e.message);
          }
        },

        async addItem(calendarId, icsData) {
          var calendar = managedCalendar(calendarId);

          var item = icsToItem(icsData);
          if (!item) {
            console.error("[ExCalSync API] addItem: ICS parse failed");
            throw new ExtensionError("Invalid ICS data");
          }

          try {
            // TB 128+ addItem returns a promise directly
            if (calendar.wrappedJSObject && typeof calendar.wrappedJSObject.addItem === "function") {
              var added = await calendar.wrappedJSObject.addItem(item);
              var resultId = added ? String(added.id || item.id) : String(item.id);
              return { id: resultId, success: true };
            }

            // Try direct promise-returning addItem
            var result = calendar.addItem(item);
            if (result && typeof result.then === "function") {
              var addedItem = await result;
              var addedId = addedItem ? String(addedItem.id || item.id) : String(item.id);
              return { id: addedId, success: true };
            }

            // Last resort: callback-based
            return await new Promise(function (resolve, reject) {
              calendar.addItem(item, {
                QueryInterface: ChromeUtils.generateQI(["calIOperationListener"]),
                onGetResult() {},
                onOperationComplete(aCal, aStatus, aOpType, aId) {
                  if (Components.isSuccessCode(aStatus)) {
                    resolve({ id: String(aId || item.id), success: true });
                  } else {
                    reject(new Error("addItem failed: " + aStatus));
                  }
                },
              });
            });
          } catch (e) {
            console.error("[ExCalSync API] addItem error:", e);
            throw new ExtensionError("addItem failed: " + e.message);
          }
        },

        async updateItem(calendarId, itemId, icsData) {
          console.log("[ExCalSync API] updateItem:", itemId);
          var calendar = managedCalendar(calendarId);

          var newItem = icsToItem(icsData);
          if (!newItem) {
            throw new ExtensionError("Invalid ICS data");
          }

          try {
            // Get all items to find the existing one
            var filter = Components.interfaces.calICalendar.ITEM_FILTER_TYPE_EVENT |
                         Components.interfaces.calICalendar.ITEM_FILTER_COMPLETED_ALL;

            var items;
            if (typeof calendar.getItemsAsArray === "function") {
              items = await calendar.getItemsAsArray(filter, 0, null, null);
            } else {
              items = await new Promise(function (resolve, reject) {
                var results = [];
                calendar.getItems(filter, 0, null, null, {
                  QueryInterface: ChromeUtils.generateQI(["calIOperationListener"]),
                  onGetResult(aCal, aStatus, aItemType, aDetail, aItems) {
                    for (var item of aItems) results.push(item);
                  },
                  onOperationComplete(aCal, aStatus) {
                    if (Components.isSuccessCode(aStatus)) resolve(results);
                    else reject(new Error("getItems failed: " + aStatus));
                  },
                });
              });
            }

            var existingItem = null;
            for (var it of items) {
              if (it.id === itemId) { existingItem = it; break; }
            }

            if (!existingItem) {
              throw new Error("Item not found: " + itemId);
            }

            var updatedItem = existingItem.clone();
            updatedItem.title = newItem.title;
            if (newItem.startDate) updatedItem.startDate = newItem.startDate;
            if (newItem.endDate) updatedItem.endDate = newItem.endDate;
            if (newItem.getProperty("DESCRIPTION")) {
              updatedItem.setProperty("DESCRIPTION", newItem.getProperty("DESCRIPTION"));
            }
            if (newItem.getProperty("LOCATION")) {
              updatedItem.setProperty("LOCATION", newItem.getProperty("LOCATION"));
            }

            // Sync the HTML description (X-ALT-DESC)
            if (newItem.getProperty("X-ALT-DESC")) {
              updatedItem.setProperty("X-ALT-DESC", newItem.getProperty("X-ALT-DESC"));
              updatedItem.setPropertyParameter("X-ALT-DESC", "FMTTYPE", "text/html");
            }

            // Sync sensitivity/class
            if (newItem.getProperty("CLASS")) {
              updatedItem.setProperty("CLASS", newItem.getProperty("CLASS"));
            }

            // Sync priority
            if (newItem.getProperty("PRIORITY")) {
              updatedItem.setProperty("PRIORITY", newItem.getProperty("PRIORITY"));
            }

            // Sync organizer
            if (newItem.organizer) {
              updatedItem.organizer = newItem.organizer;
            }

            // Sync attendees: remove old, add new
            updatedItem.removeAllAttendees();
            var newAttendees = newItem.getAttendees ? newItem.getAttendees() : [];
            for (var att of newAttendees) {
              updatedItem.addAttendee(att);
            }

            var modResult = calendar.modifyItem(updatedItem, existingItem);
            if (modResult && typeof modResult.then === "function") {
              await modResult;
            }

            return { id: String(itemId), success: true };
          } catch (e) {
            console.error("[ExCalSync API] updateItem error:", e);
            throw new ExtensionError("updateItem failed: " + e.message);
          }
        },

        async deleteItem(calendarId, itemId) {
          console.log("[ExCalSync API] deleteItem:", itemId);
          var calendar = managedCalendar(calendarId);

          try {
            var filter = Components.interfaces.calICalendar.ITEM_FILTER_TYPE_EVENT |
                         Components.interfaces.calICalendar.ITEM_FILTER_COMPLETED_ALL;

            var items;
            if (typeof calendar.getItemsAsArray === "function") {
              items = await calendar.getItemsAsArray(filter, 0, null, null);
            } else {
              items = await new Promise(function (resolve, reject) {
                var results = [];
                calendar.getItems(filter, 0, null, null, {
                  QueryInterface: ChromeUtils.generateQI(["calIOperationListener"]),
                  onGetResult(aCal, aStatus, aItemType, aDetail, aItems) {
                    for (var item of aItems) results.push(item);
                  },
                  onOperationComplete(aCal, aStatus) {
                    if (Components.isSuccessCode(aStatus)) resolve(results);
                    else reject(new Error("getItems failed: " + aStatus));
                  },
                });
              });
            }

            var targetItem = null;
            for (var it of items) {
              if (it.id === itemId) { targetItem = it; break; }
            }

            if (!targetItem) {
              throw new Error("Item not found: " + itemId);
            }

            var delResult = calendar.deleteItem(targetItem);
            if (delResult && typeof delResult.then === "function") {
              await delResult;
            }

            return { id: String(itemId), success: true };
          } catch (e) {
            console.error("[ExCalSync API] deleteItem error:", e);
            throw new ExtensionError("deleteItem failed: " + e.message);
          }
        },

        async setCalendarProperty(calendarId, key, value) {
          var calendar = managedCalendar(calendarId);
          calendar.setProperty(key, value);
          return true;
        },

        async getCalendarProperty(calendarId, key) {
          var calendar = managedCalendar(calendarId);
          var val = calendar.getProperty(key);
          return val != null ? String(val) : null;
        },

        onItemChanged: new ExtensionCommon.EventManager({
          context,
          name: "CalendarAPI.onItemChanged",
          register(fire) {
            changeObserver.listeners.add(fire);
            changeObserver.register();
            return function () {
              changeObserver.listeners.delete(fire);
              changeObserver.unregister();
            };
          },
        }).api(),
      },
    };
  }
};
