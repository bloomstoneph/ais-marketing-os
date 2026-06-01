/**
 * AIS Marketing OS — Google Apps Script
 * ─────────────────────────────────────
 * HOW TO DEPLOY:
 * 1. Go to https://script.google.com
 * 2. Create a new project → paste this entire file
 * 3. Click Deploy → New Deployment → Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the Web App URL → paste it in the app under Admin → Settings
 * 5. Click "Test Connection" to confirm it works
 *
 * SHEETS CREATED AUTOMATICALLY on first run (doGet):
 * Tasks, Subtasks, Task_Comments, Campaigns, Leave_Requests, Missions,
 * Comp_Days, Events, Shoots, Media_Assets, Content, Meetings,
 * Meeting_Actions, Departments, Members, Enrollment, Social_Metrics,
 * iCal_Feeds, Member_Goals
 */

const SHEET_NAMES = [
  'Tasks', 'Subtasks', 'Task_Comments', 'Campaigns',
  'Leave_Requests', 'Missions', 'Comp_Days', 'Events',
  'Shoots', 'Media_Assets', 'Content', 'Meetings',
  'Meeting_Actions', 'Departments', 'Members', 'Enrollment',
  'Social_Metrics', 'iCal_Feeds', 'Member_Goals',
];

// ─── GET — pull all data (called every 5s by the app) ───────────────────────
function doGet(e) {
  const params = e?.parameter || {};

  // iCal proxy — fetch external calendar and return raw ICS text
  if (params.ical) {
    try {
      const res = UrlFetchApp.fetch(params.ical, { muteHttpExceptions: true });
      return ContentService.createTextOutput(res.getContentText())
        .setMimeType(ContentService.MimeType.TEXT);
    } catch (err) {
      return ContentService.createTextOutput('ERROR:' + err.message)
        .setMimeType(ContentService.MimeType.TEXT);
    }
  }

  // Normal data pull
  try {
    const ss   = getOrCreateSpreadsheet();
    const data = {};
    SHEET_NAMES.forEach(name => {
      data[name] = sheetToArray(ss, name);
    });
    return json({ ok: true, data });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

// ─── POST — write data (append / update / delete / batch) ───────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss      = getOrCreateSpreadsheet();

    if (payload.action === 'batch') {
      // Batch: run multiple operations in one request (used by seed)
      payload.operations.forEach(op => runOp(ss, op));
      return json({ ok: true });
    }

    runOp(ss, payload);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

function runOp(ss, op) {
  const { action, sheet, id, data } = op;
  if (action === 'append') appendRow(ss, sheet, data);
  else if (action === 'update') updateRow(ss, sheet, id, data);
  else if (action === 'delete') deleteRow(ss, sheet, id);
}

// ─── SHEET HELPERS ───────────────────────────────────────────────────────────

function getOrCreateSpreadsheet() {
  // Try to find an existing spreadsheet named "AIS Marketing OS"
  const files = DriveApp.getFilesByName('AIS Marketing OS');
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  // Create a new one
  const ss = SpreadsheetApp.create('AIS Marketing OS');
  SHEET_NAMES.forEach(name => {
    try { ss.insertSheet(name); } catch(e) {}
  });
  // Remove the default "Sheet1" if it exists
  const def = ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);
  return ss;
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function sheetToArray(ss, name) {
  const sheet = getOrCreateSheet(ss, name);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] === undefined ? '' : String(row[i]); });
    return obj;
  });
}

function appendRow(ss, sheetName, data) {
  const sheet   = getOrCreateSheet(ss, sheetName);
  const allData = sheet.getDataRange().getValues();

  if (allData.length === 0 || (allData.length === 1 && allData[0].join('') === '')) {
    // Sheet is empty — write headers first
    const headers = Object.keys(data);
    sheet.appendRow(headers);
    sheet.appendRow(headers.map(h => data[h] ?? ''));
    return;
  }

  const headers = allData[0];
  // Add any new columns that don't exist yet
  const existingSet = new Set(headers);
  Object.keys(data).forEach(k => {
    if (!existingSet.has(k)) {
      headers.push(k);
      existingSet.add(k);
      const col = headers.length;
      sheet.getRange(1, col).setValue(k);
    }
  });
  sheet.appendRow(headers.map(h => data[h] ?? ''));
}

function updateRow(ss, sheetName, id, data) {
  const sheet   = getOrCreateSheet(ss, sheetName);
  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) return;

  const headers = allData[0];
  const idCol   = headers.indexOf('id');
  if (idCol < 0) return;

  // Add missing columns
  Object.keys(data).forEach(k => {
    if (!headers.includes(k)) {
      headers.push(k);
      sheet.getRange(1, headers.length).setValue(k);
    }
  });

  for (let r = 1; r < allData.length; r++) {
    if (String(allData[r][idCol]) === String(id)) {
      headers.forEach((h, c) => {
        if (data[h] !== undefined) {
          sheet.getRange(r + 1, c + 1).setValue(data[h]);
        }
      });
      return;
    }
  }
}

function deleteRow(ss, sheetName, id) {
  const sheet   = getOrCreateSheet(ss, sheetName);
  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) return;

  const headers = allData[0];
  const idCol   = headers.indexOf('id');
  if (idCol < 0) return;

  for (let r = allData.length - 1; r >= 1; r--) {
    if (String(allData[r][idCol]) === String(id)) {
      sheet.deleteRow(r + 1);
      return;
    }
  }
}

// ─── UTILS ───────────────────────────────────────────────────────────────────

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
