/**
 * Google Apps Script — append bookings to your Sheet from the dress rental website.
 *
 * SETUP (once):
 * 1. Open your bookings spreadsheet.
 * 2. Extensions → Apps Script.
 * 3. Replace the default code with this file; Save.
 * 4. Deploy → New deployment → type “Web app”:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Deploy → Authorize → copy the Web app URL (ends with /exec).
 * 6. Paste that URL into the site field “Apps Script URL (save new bookings)”.
 *
 * Row 1 must be headers (code, from, to, name). New rows append after the last row.
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!e.postData || !e.postData.contents) {
      return jsonOut(false, "No POST body.");
    }
    var data = JSON.parse(e.postData.contents);
    var code = String(data.code || "").trim();
    var from = String(data.from || "").trim();
    var to = String(data.to || "").trim();
    var name = String(data.name || "").trim();
    if (!code || !from || !to || !name) {
      return jsonOut(false, "Missing code, from, to, or name.");
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) {
      return jsonOut(false, "Dates must be YYYY-MM-DD.");
    }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    sheet.appendRow([code, from, to, name]);
    return jsonOut(true, "");
  } catch (err) {
    return jsonOut(false, String(err && err.message ? err.message : err));
  } finally {
    lock.releaseLock();
  }
}

function jsonOut(ok, errorMessage) {
  var payload = ok ? { ok: true } : { ok: false, error: errorMessage };
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON
  );
}
