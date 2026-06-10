/**
 * ─────────────────────────────────────────────────────────────────────────────
 *  SCHOOL EMERGENCY / ASSISTANCE ALERT SYSTEM — Code.gs
 *  Container-bound Google Apps Script (bound to the alerts spreadsheet).
 * ─────────────────────────────────────────────────────────────────────────────
 *
 *  DEPLOYMENT SETTINGS — these exact values are REQUIRED:
 *
 *    Deploy → New deployment → type: Web app
 *      • Execute as:      User accessing the web app
 *      • Who has access:  Anyone within [your Workspace domain]
 *
 *    Deploying "as user accessing" within the domain is what makes
 *    Session.getActiveUser().getEmail() reliably return the signed-in
 *    teacher's email, which we record in "Raised By" and "Responder".
 *    If you deploy as "Me" or allow anonymous access, the email may be blank
 *    and the audit trail is lost.
 *
 *  OAUTH SCOPES (declared in appsscript.json; users authorise on first visit):
 *      • https://www.googleapis.com/auth/spreadsheets.currentonly
 *          read/write the bound spreadsheet only
 *      • https://www.googleapis.com/auth/userinfo.email
 *          identify the submitting / responding user
 *      • https://www.googleapis.com/auth/script.storage
 *          UserProperties — remembers each teacher's last-used location
 *
 *  FIRST-TIME SETUP: run setup() once from the editor (it is idempotent and
 *  safe to re-run), then create the web app deployment as described above.
 * ─────────────────────────────────────────────────────────────────────────────
 */

/* ============================== CONSTANTS ================================= */

var SHEET_DATA = 'Data';
var SHEET_STAFF = 'Staff';
var SHEET_STUDENT = 'Student';

var TIMEZONE = 'Europe/London';

var DATA_HEADERS = [
  'ID', 'Timestamp', 'Raised By', 'Student Name', 'Location',
  'Category', 'Status', 'Responder', 'Time Responded', 'Time Resolved'
];
var STAFF_HEADERS = ['Name', 'Email', 'Role'];
var STUDENT_HEADERS = ['Name', 'ID', 'Year'];

var CATEGORIES = [
  'Urgent',
  'Medical',
  'Non-Urgent (e.g. Toilet Break)',
  'Truancy',
  'Mobile Phone'
];
var STATUSES = ['Pending', 'On Way', 'Resolved'];

// Legal status transitions, including the "undo" directions.
// On Way → Pending  (undo acknowledge: clears Responder + Time Responded)
// Resolved → On Way (undo resolve:     clears Time Resolved)
var LEGAL_TRANSITIONS = {
  'Pending': ['On Way'],
  'On Way': ['Resolved', 'Pending'],
  'Resolved': ['On Way']
};

// 1-based column indexes in the Data sheet (kept in sync with DATA_HEADERS).
var COL = {
  ID: 1, TIMESTAMP: 2, RAISED_BY: 3, STUDENT: 4, LOCATION: 5,
  CATEGORY: 6, STATUS: 7, RESPONDER: 8, TIME_RESPONDED: 9, TIME_RESOLVED: 10
};

var CACHE_TTL_SECONDS = 300; // 5 minutes for Staff/Student reference data

/* ================================ WEB APP ================================= */

/**
 * Serves the single-page app.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('School Alert System')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ================================= SETUP ================================== */

/**
 * Idempotent one-time setup. Creates the three tabs if missing, writes the
 * exact headers to row 1, freezes/bolds the header row, sets column widths,
 * applies data validation to the Status and Category columns, and sets the
 * spreadsheet timezone. Safe to run repeatedly — never wipes or duplicates
 * existing data.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TIMEZONE);

  var dataSheet = ensureSheet_(ss, SHEET_DATA, DATA_HEADERS, [
    280, 160, 220, 180, 160, 220, 100, 220, 160, 160
  ]);
  ensureSheet_(ss, SHEET_STAFF, STAFF_HEADERS, [180, 240, 160]);
  ensureSheet_(ss, SHEET_STUDENT, STUDENT_HEADERS, [180, 120, 80]);

  // Data validation on Category (F) and Status (G) for all data rows.
  var maxRows = dataSheet.getMaxRows();
  var categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CATEGORIES, true)
    .setAllowInvalid(false)
    .build();
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUSES, true)
    .setAllowInvalid(false)
    .build();
  dataSheet.getRange(2, COL.CATEGORY, maxRows - 1, 1).setDataValidation(categoryRule);
  dataSheet.getRange(2, COL.STATUS, maxRows - 1, 1).setDataValidation(statusRule);

  // Readable timestamp formatting.
  dataSheet.getRange(2, COL.TIMESTAMP, maxRows - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  dataSheet.getRange(2, COL.TIME_RESPONDED, maxRows - 1, 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');

  return 'Setup complete. Tabs verified: Data, Staff, Student.';
}

/**
 * Creates a sheet if missing and (re)applies header text + formatting without
 * touching any rows below the header.
 */
function ensureSheet_(ss, name, headers, widths) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#eff6ff'); // subtle blue header tint
  sheet.setFrozenRows(1);

  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
  return sheet;
}

/* ============================ REFERENCE DATA ============================== */

/**
 * Students for the form's predictive datalist. Cached for 5 minutes.
 * @return {{success:boolean, message:string, data:Array<{name:string,id:string,year:string}>}}
 */
function getStudents() {
  return safeCall_(function () {
    return cachedSheetRead_('students_v1', SHEET_STUDENT, function (row) {
      return { name: String(row[0]), id: String(row[1]), year: String(row[2]) };
    });
  });
}

/**
 * Staff reference data (not used for access control). Cached for 5 minutes.
 * @return {{success:boolean, message:string, data:Array<{name:string,email:string,role:string}>}}
 */
function getStaff() {
  return safeCall_(function () {
    return cachedSheetRead_('staff_v1', SHEET_STAFF, function (row) {
      return { name: String(row[0]), email: String(row[1]), role: String(row[2]) };
    });
  });
}

/**
 * Reads a reference sheet through CacheService.
 */
function cachedSheetRead_(cacheKey, sheetName, mapRow) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" not found. Run setup().');

  var lastRow = sheet.getLastRow();
  var items = [];
  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    items = values
      .filter(function (row) { return String(row[0]).trim() !== ''; })
      .map(mapRow);
  }
  cache.put(cacheKey, JSON.stringify(items), CACHE_TTL_SECONDS);
  return items;
}

/**
 * Single startup payload for the SPA: identity, last-used location and the
 * student datalist in one round trip. Every google.script.run call carries
 * ~0.5–1s of Apps Script overhead regardless of payload size, so the form
 * boots with exactly one call instead of two.
 */
function getInitData() {
  return safeCall_(function () {
    return {
      email: Session.getActiveUser().getEmail(),
      lastLocation: PropertiesService.getUserProperties().getProperty('lastLocation') || '',
      students: cachedSheetRead_('students_v1', SHEET_STUDENT, function (row) {
        return { name: String(row[0]), id: String(row[1]), year: String(row[2]) };
      })
    };
  });
}

/* ================================ ALERTS ================================== */

/**
 * Creates a new alert.
 * @param {{studentName:string, location:string, category:string}} payload
 * @return {{success:boolean, message:string, data:Object}} the created alert
 */
function createAlert(payload) {
  return safeCall_(function () {
    if (!payload || typeof payload !== 'object') throw new Error('Invalid request.');

    var studentName = String(payload.studentName || '').trim();
    var location = String(payload.location || '').trim();
    var category = String(payload.category || '').trim();

    if (!studentName) throw new Error('Student name is required.');
    if (!location) throw new Error('Location is required.');
    if (CATEGORIES.indexOf(category) === -1) throw new Error('Invalid category.');

    var sheet = dataSheet_();
    var id = Utilities.getUuid();
    var now = new Date();
    var raisedBy = Session.getActiveUser().getEmail();

    sheet.appendRow([
      id, now, raisedBy, studentName, location, category, 'Pending', '', '', ''
    ]);

    // Remember this teacher's location so the form can pre-fill next time.
    PropertiesService.getUserProperties().setProperty('lastLocation', location);

    return {
      id: id,
      timestamp: now.toISOString(),
      raisedBy: raisedBy,
      studentName: studentName,
      location: location,
      category: category,
      status: 'Pending',
      responder: '',
      timeResponded: '',
      timeResolved: ''
    };
  });
}

/**
 * Transitions an alert between statuses, including undo directions.
 * Wrapped in a script lock so two responders acting simultaneously cannot
 * corrupt a row. Returns the updated alert so the client can reconcile.
 *
 * @param {string} id        alert UUID
 * @param {string} newStatus one of 'Pending' | 'On Way' | 'Resolved'
 */
function updateAlertStatus(id, newStatus) {
  return safeCall_(function () {
    id = String(id || '').trim();
    newStatus = String(newStatus || '').trim();
    if (!id) throw new Error('Missing alert ID.');
    if (STATUSES.indexOf(newStatus) === -1) throw new Error('Invalid status.');

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      throw new Error('The system is busy — please try again.');
    }

    try {
      var sheet = dataSheet_();
      var rowIndex = findAlertRow_(sheet, id);
      if (rowIndex === -1) throw new Error('Alert not found. It may have been removed.');

      var rowRange = sheet.getRange(rowIndex, 1, 1, DATA_HEADERS.length);
      var row = rowRange.getValues()[0];
      var currentStatus = String(row[COL.STATUS - 1]);

      if (currentStatus === newStatus) {
        return rowToAlert_(row); // No-op (e.g. double-click) — return current state.
      }
      if ((LEGAL_TRANSITIONS[currentStatus] || []).indexOf(newStatus) === -1) {
        throw new Error('Cannot move an alert from "' + currentStatus + '" to "' + newStatus + '".');
      }

      var now = new Date();
      var user = Session.getActiveUser().getEmail();

      if (newStatus === 'On Way' && currentStatus === 'Pending') {
        // Acknowledge: claim the alert.
        row[COL.RESPONDER - 1] = user;
        row[COL.TIME_RESPONDED - 1] = now;
      } else if (newStatus === 'Resolved') {
        row[COL.TIME_RESOLVED - 1] = now;
      } else if (newStatus === 'Pending') {
        // Undo acknowledge: release the claim.
        row[COL.RESPONDER - 1] = '';
        row[COL.TIME_RESPONDED - 1] = '';
      } else if (newStatus === 'On Way' && currentStatus === 'Resolved') {
        // Undo resolve: reopen, keeping the original responder.
        row[COL.TIME_RESOLVED - 1] = '';
      }

      row[COL.STATUS - 1] = newStatus;
      rowRange.setValues([row]);
      SpreadsheetApp.flush();

      return rowToAlert_(row);
    } finally {
      lock.releaseLock();
    }
  });
}

/**
 * Returns only what the dashboard needs: every Pending / On Way alert, plus
 * Resolved alerts whose resolution time falls on the current calendar day in
 * the spreadsheet timezone. Older resolved alerts stay in the sheet but are
 * never sent to the client, keeping the payload small.
 */
function getActiveAlerts() {
  return safeCall_(function () {
    var sheet = dataSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var values = sheet.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();
    var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');

    var alerts = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (!row[COL.ID - 1]) continue;

      var status = String(row[COL.STATUS - 1]);
      if (status === 'Pending' || status === 'On Way') {
        alerts.push(rowToAlert_(row));
      } else if (status === 'Resolved') {
        var resolvedAt = row[COL.TIME_RESOLVED - 1];
        if (resolvedAt instanceof Date &&
            Utilities.formatDate(resolvedAt, TIMEZONE, 'yyyy-MM-dd') === today) {
          alerts.push(rowToAlert_(row));
        }
      }
    }
    return alerts;
  });
}

/* =============================== HELPERS ================================== */

function dataSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATA);
  if (!sheet) throw new Error('Sheet "' + SHEET_DATA + '" not found. Run setup().');
  return sheet;
}

/**
 * Finds the 1-based row index of an alert by ID (column A), or -1.
 * TextFinder keeps this fast even with thousands of historical rows.
 */
function findAlertRow_(sheet, id) {
  var match = sheet.getRange(2, COL.ID, Math.max(sheet.getLastRow() - 1, 1), 1)
    .createTextFinder(id)
    .matchEntireCell(true)
    .findNext();
  return match ? match.getRow() : -1;
}

/**
 * Maps a raw sheet row to the plain object sent to the client.
 * Dates are serialised as ISO strings (Date objects do not survive
 * google.script.run reliably).
 */
function rowToAlert_(row) {
  return {
    id: String(row[COL.ID - 1]),
    timestamp: toIso_(row[COL.TIMESTAMP - 1]),
    raisedBy: String(row[COL.RAISED_BY - 1] || ''),
    studentName: String(row[COL.STUDENT - 1] || ''),
    location: String(row[COL.LOCATION - 1] || ''),
    category: String(row[COL.CATEGORY - 1] || ''),
    status: String(row[COL.STATUS - 1] || ''),
    responder: String(row[COL.RESPONDER - 1] || ''),
    timeResponded: toIso_(row[COL.TIME_RESPONDED - 1]),
    timeResolved: toIso_(row[COL.TIME_RESOLVED - 1])
  };
}

function toIso_(value) {
  return (value instanceof Date) ? value.toISOString() : '';
}

/**
 * Wraps every public entry point so the client always receives a structured
 * { success, message, data } object instead of a raw thrown error.
 */
function safeCall_(fn) {
  try {
    return { success: true, message: 'OK', data: fn() };
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return { success: false, message: (err && err.message) || 'Unexpected error.', data: null };
  }
}
