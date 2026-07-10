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
  RESUME_CODE:   25   // JSON blob — paste into crawl panel to reload settings
};

const LOG_EXT_HEADERS = [
  'Type', 'Name / Sheet', 'Seeds',
  'Depth', 'Max Papers', 'Filter Groups',
  'Backward Pass', 'Status', 'Resume Code'
];

// ============================================================
// Sheet access
// ============================================================

function getLogSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
}

// Writes extended headers to row 1 if they haven't been added yet.
function ensureLogExtHeaders() {
  var sheet = getLogSheet();
  if (!sheet) return;
  // Check the last extended column — if it already has the header, all columns are set.
  if (sheet.getRange(1, LOG_EXT.RESUME_CODE).getValue() === 'Resume Code') return;
  sheet.getRange(1, LOG_EXT.TYPE, 1, LOG_EXT_HEADERS.length)
    .setValues([LOG_EXT_HEADERS])
    .setFontWeight('bold')
    .setBackground('#e8f0fe')
    .setFontColor('#1a73e8');
  // Keep the Resume Code column narrow — users copy from it, not read it at a glance.
  sheet.setColumnWidth(LOG_EXT.RESUME_CODE, 200);
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
    resumeCode
  ]]);

  setLogStatusStyle_(sheet, lastRow, 'Running');
  return lastRow;
}

// Updates the Status cell and its colour for a given log row number.
function updateLogRow(rowNum, status) {
  var sheet = getLogSheet();
  if (!sheet || !rowNum) return;
  sheet.getRange(rowNum, LOG_EXT.STATUS).setValue(status);
  setLogStatusStyle_(sheet, rowNum, status);
}

function setLogStatusStyle_(sheet, rowNum, status) {
  var cell = sheet.getRange(rowNum, LOG_EXT.STATUS);
  if (status === 'Complete') {
    cell.setBackground('#34a853').setFontColor('#ffffff');
  } else if (status.indexOf('Error') !== -1) {
    cell.setBackground('#e53935').setFontColor('#ffffff');
  } else if (status === 'Paper Limit') {
    cell.setBackground('#ff9800').setFontColor('#ffffff');
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
    var crawlSheet   = ss.getSheetByName(name);
    var isIncomplete = crawlSheet && status !== 'Complete' && status !== 'Error';

    if (isIncomplete) {
      // Restore all crawl script properties so the Resume button works correctly
      props.setProperty('CRAWL_ACTIVE_SHEET',    name);
      props.setProperty('CRAWL_DIRECTION',       'forward');
      props.setProperty('CRAWL_MAX_DEPTH',       String(depth      || 2));
      props.setProperty('CRAWL_MAX_PAPERS',      String(maxPapers  || 300));
      props.setProperty('SNOWBALL_FILTER_GROUPS', JSON.stringify(filterGroups));
      props.setProperty('CRAWL_RUN_BACKWARD',    backwardPass !== 'None' ? 'true' : 'false');
      props.setProperty('CRAWL_EXPAND_BACKWARD', backwardPass === 'Backward + Expand' ? 'true' : 'false');
      // Keep the same log row so status updates continue on the original entry
      props.setProperty('CRAWL_LOG_ROW',         String(row));
      canResume = true;
    }
  }

  var replay = {
    type:         type,
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

  if      (type === 'Crawl')    showCrawlbar();
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
