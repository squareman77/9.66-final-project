/**
 * Receives one participant's data from the drawing study and stores it twice:
 *   1. The complete JSON file in a Google Drive folder (the backup you analyze from).
 *   2. One row per drawing in the "drawings" sheet (for a quick look at incoming data).
 *
 * Setup: paste into Extensions > Apps Script of a new Google Sheet, then
 * Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 */

const FOLDER_NAME = 'drawing-study-data';
const SHEET_NAME = 'drawings';
const HEADER = ['received_at', 'participant_id', 'block', 'shape', 'trial_index',
                'position', 'vertices', 'canvas_size', 'rt_ms', 'n_clears', 'n_undos'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // participants finishing at the same moment take turns
  try {
    const raw = e.postData.contents;
    const data = JSON.parse(raw);
    const pid = String(data.participant_id || 'unknown').replace(/[^\w-]/g, '_');

    // 1. Full backup file
    getFolder_().createFile(pid + '_' + Date.now() + '.json', raw, MimeType.PLAIN_TEXT);

    // 2. One row per drawing
    const now = new Date();
    const rows = [];
    (data.trials || []).forEach(function (t) {
      const items = t.block === 'teach' ? t.items : [t];
      items.forEach(function (it) {
        rows.push([now, data.participant_id, t.block, t.shape, t.trial_index,
                   t.block === 'teach' ? it.position : '',
                   JSON.stringify(it.vertices), it.canvas_size,
                   t.block === 'teach' ? t.rt_ms : it.rt_ms, it.n_clears, it.n_undos]);
      });
    });
    const sheet = getSheet_();
    if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADER.length).setValues(rows);

    return json_({ ok: true, rows: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Visiting the web-app URL in a browser shows this, so you can check the deployment works.
function doGet() {
  return json_({ ok: true, message: 'Drawing study endpoint is running.' });
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADER);
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}