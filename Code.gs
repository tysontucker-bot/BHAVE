/**
 * BHAVE — behavior logging for aides. Google Apps Script backend.
 *
 * Setup (one time):
 *   1. In your Google Sheet: Extensions → Apps Script. Delete the sample code, paste ALL of this file.
 *   2. Save. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone → Deploy.
 *   3. Copy the Web app URL (ends in /exec) and send it to Claude.
 *
 * Students and behaviors: edit the "Setup" tab (created automatically the first time the app opens).
 *   One row per behavior:  Initials | Behavior | Color (optional, like #1E88E5)
 * Staff names: edit the "Staff" tab, one name per row. These are the "Who's using this phone?" choices.
 * Every logged entry goes to the "Log" tab.
 *
 * Note: Google reserves the URL parameters "c" and "sid", so this script never uses them.
 */
const SECRET_CODE = "y7nkjobi14jz78w5";

const SETUP_TAB = "Setup";
const LOG_TAB = "Log";
const SETUP_HEADERS = ["Initials", "Behavior", "Color (optional)"];
const LOG_HEADERS = ["Date", "Time", "Student", "Behavior", "Amount", "Type", "Logged by", "Entered at"];
const STAFF_TAB = "Staff";
const STAFF_HEADERS = ["Name"];
const DEFAULT_STAFF = ["Para A", "Para B", "Teacher"];
const PALETTE = ["#1E88E5", "#8E24AA", "#F4511E", "#00897B", "#D81B60", "#6D4C41", "#3949AB", "#7CB342", "#F9A825", "#5E35B1"];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.key !== SECRET_CODE) return json({ ok: false, error: "wrong-code" });

  if (p.action === "setup") return json({ ok: true, students: readSetup(), staff: readStaff() });

  if (p.action === "log") {
    let items;
    try { items = JSON.parse(p.items || "[]"); } catch (err) { return json({ ok: false, error: "bad-data" }); }
    if (!Array.isArray(items) || !items.length || items.length > 50) return json({ ok: false, error: "bad-data" });
    // Only names on the Staff tab are recorded (anything else is left blank)
    const by = readStaff().indexOf(String(p.by || "")) !== -1 ? String(p.by) : "";
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const now = new Date();
    // The phone saves entries instantly and may send them a little later (or the next time the app opens),
    // so it sends when Submit was tapped ("ent"). Date and "Entered at" come from that, not from arrival time.
    let entered = now;
    const ent = Number(p.ent);
    if (ent && ent <= now.getTime() + 5 * 60000 && ent >= now.getTime() - 7 * 86400000) entered = new Date(ent);
    const date = Utilities.formatDate(entered, tz, "M/d/yyyy");
    const enteredAt = Utilities.formatDate(entered, tz, "M/d/yyyy h:mm a");
    // "Time" is when it happened: the aide can set an earlier time that day (HH:mm); otherwise it's when entered
    let time = Utilities.formatDate(entered, tz, "h:mm a");
    const at = String(p.at || "").match(/^(\d{1,2}):(\d{2})$/);
    if (at) {
      const h = Number(at[1]), m = Number(at[2]);
      const enteredMin = Number(Utilities.formatDate(entered, tz, "H")) * 60 + Number(Utilities.formatDate(entered, tz, "m"));
      if (h > 23 || m > 59) return json({ ok: false, error: "bad-time" });
      // A few minutes of leeway so a phone clock that's slightly off can't block a save
      if (h * 60 + m > enteredMin + 5) return json({ ok: false, error: "future-time" });
      time = (h % 12 || 12) + ":" + (m < 10 ? "0" + m : m) + (h < 12 ? " AM" : " PM");
    }
    // Each submission has a unique id: if the phone retries one that already arrived, don't save it twice
    const id = String(p.id || "").slice(0, 40);
    const cache = CacheService.getScriptCache();
    if (id && cache.get("sub_" + id)) return json({ ok: true, saved: 0, duplicate: true });
    const rows = [];
    items.forEach(function (it) {
      const amount = Math.round(Number(it.n));
      const type = it.t === "min" ? "minutes" : "count";
      if (!it.s || !it.b || !(amount > 0) || amount > 999) return;
      rows.push({
        "Date": date, "Time": time, "Student": String(it.s).slice(0, 10), "Behavior": String(it.b).slice(0, 60),
        "Amount": amount, "Type": type, "Logged by": by, "Entered at": enteredAt,
      });
    });
    if (!rows.length) return json({ ok: false, error: "bad-data" });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getTab(LOG_TAB, LOG_HEADERS);
      // Write by column name, adding any missing columns (like "Entered at") at the end,
      // so older Log tabs keep their rows lined up
      let headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getDisplayValues()[0];
      LOG_HEADERS.forEach(function (h) {
        if (headers.indexOf(h) === -1) {
          sheet.getRange(1, headers.length + 1).setValue(h).setFontWeight("bold");
          headers = headers.concat([h]);
        }
      });
      const out = rows.map(function (r) { return headers.map(function (h) { return h in r ? r[h] : ""; }); });
      if (id && cache.get("sub_" + id)) return json({ ok: true, saved: 0, duplicate: true }); // arrived twice at once
      sheet.getRange(sheet.getLastRow() + 1, 1, out.length, headers.length).setValues(out);
      if (id) cache.put("sub_" + id, "1", 21600); // remembered for 6 hours
    } finally {
      lock.releaseLock();
    }
    return json({ ok: true, saved: rows.length });
  }

  return json({ ok: false, error: "unknown-action" });
}

// Reads the Setup tab into [{ initials, color, behaviors: [...] }], in the order rows appear
function readSetup() {
  const sheet = getTab(SETUP_TAB, SETUP_HEADERS);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const values = sheet.getRange(2, 1, last - 1, 3).getDisplayValues();
  const byInitials = {};
  const order = [];
  values.forEach(function (r) {
    const initials = String(r[0]).trim().toUpperCase();
    const behavior = String(r[1]).trim();
    const color = String(r[2]).trim();
    if (!initials || !behavior) return;
    if (!byInitials[initials]) {
      byInitials[initials] = { initials: initials, color: "", behaviors: [] };
      order.push(initials);
    }
    const st = byInitials[initials];
    if (!st.color && /^#[0-9a-fA-F]{6}$/.test(color)) st.color = color;
    if (st.behaviors.indexOf(behavior) === -1) st.behaviors.push(behavior);
  });
  return order.map(function (k, i) {
    const st = byInitials[k];
    if (!st.color) st.color = PALETTE[i % PALETTE.length];
    return st;
  });
}

// Reads the Staff tab: one name per row, in order
function readStaff() {
  const sheet = getTab(STAFF_TAB, STAFF_HEADERS);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const names = [];
  sheet.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(function (r) {
    const name = String(r[0]).trim().slice(0, 30);
    if (name && names.indexOf(name) === -1) names.push(name);
  });
  return names;
}

function getTab(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    if (name === SETUP_TAB) {
      // Example rows so the format is obvious; replace them with your students
      sheet.getRange(2, 1, 2, 3).setValues([["EX", "Example behavior", ""], ["EX", "Another behavior", ""]]);
    }
    if (name === STAFF_TAB) {
      sheet.getRange(2, 1, DEFAULT_STAFF.length, 1).setValues(DEFAULT_STAFF.map(function (n) { return [n]; }));
    }
  }
  return sheet;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
