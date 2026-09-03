/**
 *  Wedding RSVP + Wishes — Google Apps Script
 *
 *  Setup (about 10 minutes, once):
 *  1. Create a new Google Sheet. Name it anything.
 *  2. Extensions → Apps Script. Delete everything there, paste this file.
 *  3. Click Deploy → New deployment.
 *       Type:          Web app
 *       Execute as:    Me
 *       Who has access: Anyone            ← must be "Anyone", not "Anyone with Google account"
 *     Click Deploy, authorise when asked, copy the Web app URL.
 *  4. Paste that URL into CONFIG.endpoint in index.html.
 *
 *  If you ever edit this script, you must Deploy → Manage deployments →
 *  edit → New version, or the live URL keeps running the old code.
 */

const RSVP_SHEET = "RSVP";
const WISH_SHEET = "Wishes";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    if (data.kind === "wish") {
      const sh = sheet(ss, WISH_SHEET, ["Time", "Name", "Wish"]);
      sh.appendRow([new Date(), clean(data.name, 80), clean(data.message, 500)]);
    } else {
      const sh = sheet(ss, RSVP_SHEET, ["Time", "Name", "Attending", "Guests", "Note"]);
      sh.appendRow([
        new Date(),
        clean(data.name, 80),
        data.attend === "Yes" ? "Yes" : "No",
        Number(data.pax) || 0,
        clean(data.note, 500)
      ]);
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === "wishes") {
    const n  = Math.min(Number(p.n) || 3, 20);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(WISH_SHEET);
    let wishes = [];
    if (sh && sh.getLastRow() > 1) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
      wishes = rows.slice(-n).reverse().map(r => ({ name: r[1], message: r[2] }));
    }
    return json({ wishes });
  }
  return json({ ok: true, hint: "POST an RSVP or wish, or GET ?action=wishes" });
}

/* ---- helpers ---- */
function sheet(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
function clean(v, max) { return String(v || "").trim().slice(0, max); }
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
