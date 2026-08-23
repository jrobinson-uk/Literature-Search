// ============================================================
// Log — shared utilities for the "log" sheet
// Appends rows for Scholar, Snowball and Crawl processes and
// supports resuming / replaying any process from any log row.
// ============================================================

const LOG_SHEET_NAME = 'log';
const LOG_DATA_START = 3;   // rows 1–2 are existing headers; data starts row 3

// Extended columns appended to the right of the existing 16 Scholar columns.
const LOG_EXT = {
  TYPE:          17,  // Scholar | Snowball | Crawl
  NAME:          18,  // crawl sheet name / snowball seed title
  SEEDS:         19,  // pipe-separated seed IDs
  DEPTH:         20,  // crawl max depth
  MAX_PAPERS:    21,  // crawl max papers
  FILTER_GROUPS: 22,  // JSON filter groups
  BACKWARD_PASS: 23,  // None | Backward | Backward + Expand
  STATUS:        24,  // Running | Complete | Error | Paper Limit
  RESUME_CODE:   25,  // JSON blob — paste into crawl panel to reload settings
  COMPLETED_AT:  26   // set whenever status reaches a terminal state — see isTerminalStatus_
};

// Column 1 (TimeStamp, set in appendLogRow) already serves as "started at".
const LOG_EXT_HEADERS = [
  'Type', 'Name / Sheet', 'Seeds',
  'Depth', 'Max Papers', 'Filter Groups',
  'Backward Pass', 'Status', 'Resume Code', 'Completed At'
];

// ============================================================
// Sheet access
// ============================================================

// Case-insensitive lookup — getSheetByName() is case-sensitive in Apps Script,
// but the sheet is created as "Log" (Code.js) while this module refers to it
// as lowercase 'log', so an exact-case lookup would never find it.
function getLogSheet() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase() === LOG_SHEET_NAME) return sheets[i];
  }
  return null;
}

// Writes extended headers to row 1 if they haven't been added yet.
function ensureLogExtHeaders() {
  var sheet = getLogSheet();
  if (!sheet) return;
  // Check the last extended column — if it already has the header, all columns
  // are set. Checking the true last column (not an earlier one, e.g. Resume
  // Code) matters so a sheet from before a new column was added gets it
  // backfilled rather than being mistaken for already fully set up.
  if (sheet.getRange(1, LOG_EXT.COMPLETED_AT).getValue() === 'Completed At') return;
  sheet.getRange(1, LOG_EXT.TYPE, 1, LOG_EXT_HEADERS.length)
    .setValues([LOG_EXT_HEADERS])
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setFontColor('#1a73e8');
  // Keep the Resume Code column fairly narrow — users copy from it, not read it
  // at a glance — but wide enough to reduce the temptation to drag-select just
  // the visible text (which would silently truncate the JSON on copy).
  sheet.setColumnWidth(LOG_EXT.RESUME_CODE, 320);
}

// ============================================================
// Writing log rows
// ============================================================

// Appends a new log row and returns its row number.
//
// type  : 'Scholar' | 'Snowball' | 'Crawl'
// data  : {
//   name          : string   — crawl sheet name or search label
//   seeds         : string[] — array of ID strings (e.g. ["W123", "S2:abc"])
//   depth         : number   — crawl max depth (omit for Snowball/Scholar)
//   maxPapers     : number   — crawl max papers
//   filterGroups  : array    — filter group objects
//   runBackward   : bool
//   expandBackward: bool
// }
function appendLogRow(type, data) {
  var sheet = getLogSheet();
  if (!sheet) return null;

  ensureLogExtHeaders();

  var tz        = Session.getScriptTimeZone();
  var timestamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var lastRow   = Math.max(sheet.getLastRow(), LOG_DATA_START - 1) + 1;

  // Scholar rows already have col 1 written by the existing search code;
  // for Snowball and Crawl we write the timestamp ourselves.
  if (type !== 'Scholar') {
    sheet.getRange(lastRow, 1).setValue(timestamp);
  }

  var backwardLabel = 'None';
  if (data.runBackward && data.expandBackward) backwardLabel = 'Backward + Expand';
  else if (data.runBackward)                   backwardLabel = 'Backward';

  var seedStr = (data.seeds || []).join(' | ');

  // Compact JSON that can be pasted directly into the crawl / snowball panel
  // to restore all settings for a fresh run or to tweak and re-run.
  var resumeCode = JSON.stringify({
    type:         type,
    name:         data.name         || '',
    seedsStr:     seedStr,
    depth:        data.depth        != null ? data.depth    : '',
    maxPapers:    data.maxPapers    != null ? data.maxPapers : '',
    filterGroups: data.filterGroups || [],
    backwardPass: backwardLabel,
    canResume:    false   // paste-in always starts a fresh run
  });

  sheet.getRange(lastRow, LOG_EXT.TYPE, 1, LOG_EXT_HEADERS.length).setValues([[
    type,
    data.name         || '',
    seedStr,
    data.depth        != null ? data.depth    : '',
    data.maxPapers    != null ? data.maxPapers : '',
    data.filterGroups ? JSON.stringify(data.filterGroups) : '',
    backwardLabel,
    'Running',
    resumeCode,
    ''   // Completed At — filled in by updateLogRow once a terminal status is reached
  ]]);

  setLogStatusStyle_(sheet, lastRow, 'Running');
  return lastRow;
}

// True for any status that means the trigger has stopped running — as
// opposed to "Running" / in-progress batch messages like "batch 3 done,
// batch 4 starting…". Paper Limit and Shortfall (v3's keyword-pass
// shortfall stop) count as terminal even though they're resumable, since
// the crawl genuinely stops until a manual Resume.
function isTerminalStatus_(status) {
  return status === 'Complete' || status === 'Cancelled' || status === 'Paper Limit' ||
         status === 'Shortfall' || status.indexOf('Error') !== -1;
}

// Updates the Status cell and its colour for a given log row number.
// Stamps Completed At whenever the status lands on a terminal state —
// overwritten each time if a resumed run later stops again.
function updateLogRow(rowNum, status) {
  var sheet = getLogSheet();
  if (!sheet || !rowNum) return;
  sheet.getRange(rowNum, LOG_EXT.STATUS).setValue(status);
  setLogStatusStyle_(sheet, rowNum, status);
  if (isTerminalStatus_(status)) {
    var tz = Session.getScriptTimeZone();
    sheet.getRange(rowNum, LOG_EXT.COMPLETED_AT)
      .setValue(Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss'));
  }
}

function setLogStatusStyle_(sheet, rowNum, status) {
  var cell = sheet.getRange(rowNum, LOG_EXT.STATUS);
  if (status === 'Complete') {
    cell.setBackground('#34a853').setFontColor('#ffffff');
  } else if (status.indexOf('Error') !== -1) {
    cell.setBackground('#e53935').setFontColor('#ffffff');
  } else if (status === 'Paper Limit' || status === 'Shortfall') {
    cell.setBackground('#ff9800').setFontColor('#ffffff');
  } else if (status === 'Cancelled') {
    cell.setBackground('#9e9e9e').setFontColor('#ffffff');
  } else {
    // Running / transitional
    cell.setBackground('#f4b400').setFontColor('#333333');
  }
}

// ============================================================
// Resume from log
// ============================================================

// Menu handler: user selects a row in the log sheet then runs this.
// Stores replay data in script properties and opens the right panel.
function resumeFromLog() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui    = SpreadsheetApp.getUi();

  if (sheet.getName().toLowerCase() !== LOG_SHEET_NAME) {
    ui.alert('Please select a row in the "log" sheet first, then run this again.');
    return;
  }

  var row = sheet.getActiveCell().getRow();
  if (row < LOG_DATA_START) {
    ui.alert('Please select a data row (row 3 or below), not the header rows.');
    return;
  }

  var ext  = sheet.getRange(row, LOG_EXT.TYPE, 1, LOG_EXT_HEADERS.length).getValues()[0];
  var type = String(ext[0] || '').trim();

  if (!type) {
    ui.alert(
      'This row has no process type.\n\n' +
      'Only rows logged by version 1.11.5+ can be resumed.\n' +
      'Older Scholar rows do not carry resume data.'
    );
    return;
  }

  var name            = String(ext[1] || '').trim();
  var seedsStr        = String(ext[2] || '').trim();
  var depth           = ext[3];
  var maxPapers       = ext[4];
  var filterGroupsStr = String(ext[5] || '').trim();
  var backwardPass    = String(ext[6] || '').trim();
  var status          = String(ext[7] || '').trim();

  var filterGroups = [];
  try { if (filterGroupsStr) filterGroups = JSON.parse(filterGroupsStr); } catch (e) {}

  var props      = PropertiesService.getScriptProperties();
  var canResume  = false;

  if (type === 'Crawl') {
    // v1's own orchestration (crawlBatchTrigger, resumeCrawl) was retired in
    // v22 — an old v1 sheet genuinely can't be resumed mid-crawl any more,
    // so canResume stays false here rather than restoring CRAWL_* properties
    // for a trigger function that no longer exists. The seeds/filter groups
    // below still get carried into the replay, though, so the v2 panel can
    // pre-fill them for a FRESH crawl reusing the same configuration.
  } else if (type === 'CrawlV2') {
    var crawlV2Sheet   = ss.getSheetByName(name);
    var isIncompleteV2 = crawlV2Sheet && status !== 'Complete' && status.indexOf('Error') === -1;

    if (isIncompleteV2) {
      // Mirrors the Crawl branch above, but into v2's own CRAWL2_* namespace
      // — a resumed v2 run must not touch the v1 CRAWL_* properties, so a
      // v1 crawl running at the same time is unaffected.
      // Note: yearBound/yearFloor/yearCeiling/targetSeeds/phase aren't part
      // of the Log sheet's schema, so a resumed run restarts from Phase 1
      // (keyword) with year-bound defaults rather than exactly where it left
      // off — same accepted limitation as v1's resumeFromLog.
      props.setProperty('CRAWL2_ACTIVE_SHEET',  name);
      props.setProperty('CRAWL2_MAX_DEPTH',     String(depth      || 3));
      props.setProperty('CRAWL2_MAX_PAPERS',    String(maxPapers  || 300));
      props.setProperty('CRAWL2_FILTER_GROUPS', JSON.stringify(filterGroups));
      props.setProperty('CRAWL2_RUN_BACKWARD',  backwardPass !== 'None' ? 'true' : 'false');
      props.setProperty('CRAWL2_LOG_ROW',       String(row));
      canResume = true;
    }
  }

  // Old v1 'Crawl' rows now open in the v2 panel too (v1's own panel/
  // orchestration is retired), so the replay's type is presented to that
  // panel as 'CrawlV2' — the panel's own replay handler only auto-applies
  // a 'CrawlV2'-typed object. The STORED log-row type is untouched; this
  // only affects what this one resumeFromLog() call hands the panel.
  var replayType = (type === 'Crawl') ? 'CrawlV2' : type;

  var replay = {
    type:         replayType,
    name:         name,
    seedsStr:     seedsStr,        // pipe-separated IDs → panel converts to comma-sep input
    depth:        depth,
    maxPapers:    maxPapers,
    filterGroups: filterGroups,
    backwardPass: backwardPass,
    status:       status,
    logRow:       row,
    canResume:    canResume        // true = incomplete crawl, Resume button will work
  };

  props.setProperty('LOG_REPLAY', JSON.stringify(replay));

  if      (type === 'Crawl')    showCrawlV2bar(); // v1's own panel is retired — opens in the (now sole) crawl panel instead
  else if (type === 'CrawlV2')  showCrawlV2bar();
  else if (type === 'Snowball') showSnowballbar();
  else                          showSidebar();
}

// Called by panels on load — returns the pending replay object and clears it.
// Returns null if no replay is pending.
function getLogReplay() {
  var props = PropertiesService.getScriptProperties();
  var raw   = props.getProperty('LOG_REPLAY');
  if (!raw) return null;
  props.deleteProperty('LOG_REPLAY');
  try   { return JSON.parse(raw); }
  catch (e) { return null; }
}
