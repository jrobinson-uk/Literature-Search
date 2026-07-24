// ============================================================
// Crawl — resumable forward/backward citation crawl with JS filter
// Direction is set at crawl start and stored in script properties.
// The sheet itself is the queue: Crawled=FALSE rows are pending.
// ============================================================

// Cols 1-14 are "owned" data/formula columns.
// Col 15 is "In-Sheet Links" — a MAP formula column, not in CRAWL_HEADERS.
// Term-helper columns from applyCrawlHighlight start at CRAWL_FIRST_DETAIL_COL (16).
const CRAWL_HEADERS  = ["Depth","Crawled","Year","Title","Authors","Type","Venue","Abstract","ID","Cited By","Filter Match","Found From","Matched Cites","Direction"];
const CRAWL_NUM_COLS = CRAWL_HEADERS.length; // 14

const CRAWL_COL = {
  DEPTH:         1,
  CRAWLED:       2,
  YEAR:          3,
  TITLE:         4,
  AUTHORS:       5,
  PUB_TYPES:     6,
  VENUE:         7,
  ABSTRACT:      8,
  ID:            9,
  CITED_BY:      10,
  FILTER_MATCH:  11,
  FOUND_FROM:    12,  // parent paper ID (written at crawl time)
  MATCHED_CITES: 13, // in-sheet child IDs (written by updateCrawlMatchedCites)
  DIRECTION:     14  // "F" = forward pass, "B" = backward pass
};

// Col 15: In-Sheet Links — MAP formula; sits after Direction and before term helpers.
// Counts how many papers in the sheet list this paper as their Found From parent (col L).
const CRAWL_IN_SHEET_LINKS_COL = CRAWL_COL.DIRECTION + 1; // 15
const CRAWL_IN_SHEET_LINKS_FORMULA =
  '=MAP(A2:A,I2:I,LAMBDA(a,id,' +
  'IF(ROW(a)=2,"In-Sheet Links",' +
  'IF(a="","",COUNTIF(L$3:L,id)))))';

// Term-helper columns written by applyCrawlHighlight start here.
const CRAWL_FIRST_DETAIL_COL = CRAWL_IN_SHEET_LINKS_COL + 1; // 16

// Column letter for FILTER_MATCH (col 11 = K) used in CF formula
const CRAWL_FILTER_MATCH_COL_LETTER = "K";

// Default MAP formula — evaluates to FALSE until a real filter is applied
const CRAWL_DEFAULT_FILTER_FORMULA =
  '=MAP(A2:A,H2:H,LAMBDA(a,h,IF(ROW(a)=2,"Filter Match",IF(a<>"",FALSE,""))))';

const CRAWL_ROW_HEIGHT    = 60;
// 4 min, not 5 — a single loop iteration's fetch/back-off can occasionally run
// long (S2 429 retries alone can take ~43s), so this leaves more headroom
// before Apps Script's hard 6-min per-execution kill.
const CRAWL_TIME_LIMIT_MS = 4 * 60 * 1000;
// How many consecutive crawlBatchTrigger failures (any uncaught exception) are
// tolerated before giving up and deleting the trigger — below this, a failure
// is assumed transient and the still-alive trigger retries next minute.
const CRAWL_MAX_CONSEC_FAILURES = 3;

// Semantic Scholar back-off delays (ms) applied on HTTP 429 responses.
// Sequence: 3 s → 10 s → 30 s (then give up / surface error).
const S2_BACKOFF_MS = [3000, 10000, 30000];

// Row-1 status indicator — sits in the Year + Title columns of the header row,
// which are otherwise unused at row 1.
const CRAWL_STATUS_LABEL_COL = 3;  // "Crawl Status" label (Year col, row 1)
const CRAWL_STATUS_VALUE_COL = 4;  // status value with colour coding (Title col, row 1)

// ============================================================
// Semantic Scholar fetch helpers
// ============================================================

// Returns UrlFetchApp options for every S2 request.
// Includes x-api-key header when a key has been stored via the menu.
function getS2FetchOptions() {
  var key  = PropertiesService.getScriptProperties().getProperty('S2_API_KEY') || '';
  var opts = { muteHttpExceptions: true };
  if (key) opts.headers = { 'x-api-key': key };
  return opts;
}

// Fetches a Semantic Scholar URL, silently retrying on 429 with
// exponential back-off (3 s → 10 s → 30 s).  Used for background
// crawl-loop and seed-resolution calls.
function s2Fetch(url) {
  var opts = getS2FetchOptions();
  var resp;
  for (var i = 0; i <= S2_BACKOFF_MS.length; i++) {
    resp = UrlFetchApp.fetch(url, opts);
    if (resp.getResponseCode() !== 429 || i === S2_BACKOFF_MS.length) break;
    Utilities.sleep(S2_BACKOFF_MS[i]);
  }
  return resp;
}

// ============================================================
// Status cell helpers
// ============================================================

// Updates the Crawl Status value cell in row 1 with colour coding:
//   green  = Complete
//   red    = Error
//   grey   = Cancelled
//   amber  = Running / transitional
function setCrawlStatus(sheet, status) {
  var cell = sheet.getRange(1, CRAWL_STATUS_VALUE_COL);
  cell.setValue(status).setFontWeight('bold').setFontSize(10);
  if (status.indexOf('Complete') !== -1) {
    cell.setBackground('#34a853').setFontColor('#ffffff');
  } else if (status.indexOf('Error') !== -1) {
    cell.setBackground('#e53935').setFontColor('#ffffff');
  } else if (status === 'Cancelled') {
    cell.setBackground('#9e9e9e').setFontColor('#ffffff');
  } else {
    cell.setBackground('#f4b400').setFontColor('#333333');
  }
}

// ============================================================
// Trigger management
// ============================================================

// Deletes the stored crawl trigger (if any) and clears its stored ID.
function deleteCrawlTrigger() {
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty('CRAWL_TRIGGER_ID');
  if (!id) return;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getUniqueId() === id) ScriptApp.deleteTrigger(t);
  });
  props.deleteProperty('CRAWL_TRIGGER_ID');
}

// True only if the stored trigger ID actually matches a currently-registered
// trigger — CRAWL_ACTIVE_SHEET can still be set after a crawl completed,
// errored, or hit its cap, so that alone isn't a reliable "is it running" check.
function isCrawlRunning() {
  var id = PropertiesService.getScriptProperties().getProperty('CRAWL_TRIGGER_ID');
  if (!id) return false;
  return ScriptApp.getProjectTriggers().some(function(t) { return t.getUniqueId() === id; });
}

// Menu handler: stops the running crawl's trigger without wiping its state,
// so "Resume Last Crawl" still works afterward if the user changes their mind.
function cancelCrawl() {
  var ui = SpreadsheetApp.getUi();
  if (!isCrawlRunning()) {
    ui.alert('No crawl is currently running.');
    return;
  }
  var props     = PropertiesService.getScriptProperties();
  var sheetName = props.getProperty('CRAWL_ACTIVE_SHEET');
  var response  = ui.alert('Cancel Crawl',
    'Stop the currently running crawl ("' + sheetName + '")?\n\n' +
    'Progress so far is kept — you can still click "Resume Last Crawl" later.',
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  deleteCrawlTrigger();
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var logRow = parseInt(props.getProperty('CRAWL_LOG_ROW') || '0') || 0;
  if (sheet)  setCrawlStatus(sheet, 'Cancelled');
  if (logRow) updateLogRow(logRow, 'Cancelled');
  ui.alert('Crawl cancelled.');
}

// Creates a 1-minute repeating trigger pointing at crawlBatchTrigger,
// removing any stale trigger first.
function createCrawlTrigger() {
  deleteCrawlTrigger();
  var trigger = ScriptApp.newTrigger('crawlBatchTrigger')
    .timeBased().everyMinutes(1).create();
  PropertiesService.getScriptProperties()
    .setProperty('CRAWL_TRIGGER_ID', trigger.getUniqueId());
}

// ============================================================
// Background batch handler (called by the time-based trigger)
// ============================================================

// Runs one crawl batch.  A LockService guard ensures only one invocation runs at
// a time — extra trigger firings while a batch is running exit immediately.
// Handles three phases: "forward" → "backward" → (if expandBackward) "forward" again.
function crawlBatchTrigger() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return; // another batch is still running

  try {
    var props     = PropertiesService.getScriptProperties();
    // Reset up front — if this invocation throws, the catch block below
    // increments from this baseline, so only genuinely consecutive failures
    // (not this execution's own past attempts) count toward the threshold.
    props.setProperty('CRAWL_CONSEC_FAILURES', '0');
    var sheetName = props.getProperty('CRAWL_ACTIVE_SHEET');
    if (!sheetName) { deleteCrawlTrigger(); return; }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) { deleteCrawlTrigger(); return; }

    var phase          = props.getProperty('CRAWL_PHASE')          || 'forward';
    var batch          = parseInt(props.getProperty('CRAWL_BATCH_NUM') || '1');
    var direction      = props.getProperty('CRAWL_DIRECTION')       || 'forward';
    var groups         = JSON.parse(props.getProperty('SNOWBALL_FILTER_GROUPS') || '[]');
    var maxDepth       = parseInt(props.getProperty('CRAWL_MAX_DEPTH')   || '2');
    var maxPapers      = parseInt(props.getProperty('CRAWL_MAX_PAPERS')  || '300');
    var runBackward    = props.getProperty('CRAWL_RUN_BACKWARD')    === 'true';
    var expandBackward = props.getProperty('CRAWL_EXPAND_BACKWARD') === 'true';
    var matchesOnly    = props.getProperty('CRAWL_MATCHES_ONLY')    !== 'false';
    var yearFloor      = parseInt(props.getProperty('CRAWL_YEAR_FLOOR')   || '0') || 0;
    var yearCeiling    = parseInt(props.getProperty('CRAWL_YEAR_CEILING') || '0') || 0;
    var yearBound      = props.getProperty('CRAWL_YEAR_BOUND')      !== 'false';
    var backwardDepth  = parseInt(props.getProperty('CRAWL_BACKWARD_DEPTH') || '1') || 1;
    var logRow         = parseInt(props.getProperty('CRAWL_LOG_ROW') || '0') || 0;

    var result;

    // ── Forward phase ─────────────────────────────────────────
    if (phase === 'forward') {
      // paperDir is 'B' when we're expanding backward-discovered papers
      var paperDir = props.getProperty('CRAWL_EXPANDING_BACKWARD') === 'true' ? 'B' : 'F';
      var phaseLabel = paperDir === 'B' ? 'Expanding backward papers' : 'Forward crawl';

      if (countUncrawled(sheet) === 0) {
        // Nothing left to crawl — transition
        return transitionFromForward(props, sheet, runBackward, expandBackward);
      }

      setCrawlStatus(sheet, phaseLabel + ' — batch ' + batch + '…');
      result = runCrawlLoop(sheet, direction, groups, maxDepth, maxPapers, paperDir, matchesOnly, yearFloor, yearCeiling, yearBound);

      if (result.status === 'time-limit') {
        props.setProperty('CRAWL_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, phaseLabel + ' — batch ' + batch + ' done, batch ' + (batch + 1) + ' starting…');
      } else if (result.status === 'paper-limit') {
        // Hard cap reached — a deliberate stop, not completion. The log row's
        // 'Paper Limit' status was already set inside runCrawlLoop; don't run
        // it through transitionFromForward, which would overwrite it as
        // 'Complete' even though the queue still has unprocessed rows.
        deleteCrawlTrigger();
        setCrawlStatus(sheet, result.message);
      } else {
        transitionFromForward(props, sheet, runBackward, expandBackward, logRow);
      }

    // ── Backward phase ────────────────────────────────────────
    } else if (phase === 'backward') {
      setCrawlStatus(sheet, 'Backward pass — batch ' + batch + '…');
      result = runBackwardPass(sheet, groups, backwardDepth, maxPapers, expandBackward, matchesOnly, yearFloor, yearCeiling, yearBound);

      if (result.status === 'time-limit') {
        props.setProperty('CRAWL_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, 'Backward pass — batch ' + batch + ' done, batch ' + (batch + 1) + ' starting…');
      } else {
        // Backward pass complete — only ever do this once per crawl, so
        // transitionFromForward doesn't re-enter it again after the
        // expand-backward forward phase (below) itself completes.
        props.setProperty('CRAWL_BACKWARD_DONE', 'true');
        if (expandBackward && countUncrawled(sheet) > 0) {
          // Backward papers written with Crawled=FALSE — run forward on them
          props.setProperty('CRAWL_PHASE',              'forward');
          props.setProperty('CRAWL_EXPANDING_BACKWARD', 'true');
          props.setProperty('CRAWL_BATCH_NUM',          '1');
          setCrawlStatus(sheet, 'Backward pass complete — expanding backward papers…');
        } else {
          deleteCrawlTrigger();
          updateLogRow(logRow, 'Complete');
          setCrawlStatus(sheet, 'Complete');
        }
      }
    }

  } catch (e) {
    var propsE    = PropertiesService.getScriptProperties();
    var failCount = (parseInt(propsE.getProperty('CRAWL_CONSEC_FAILURES') || '0') || 0) + 1;
    propsE.setProperty('CRAWL_CONSEC_FAILURES', String(failCount));

    try {
      var es  = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName(propsE.getProperty('CRAWL_ACTIVE_SHEET'));
      var elr = parseInt(propsE.getProperty('CRAWL_LOG_ROW') || '0') || 0;

      if (failCount >= CRAWL_MAX_CONSEC_FAILURES) {
        // Persistent failure — give up and surface it clearly.
        deleteCrawlTrigger();
        var msg = 'Error: ' + e.message.slice(0, 60);
        if (es)  setCrawlStatus(es, msg);
        if (elr) updateLogRow(elr, msg);
      } else {
        // Likely transient (a momentary API/Sheets hiccup) — leave the
        // trigger alive so the next minute's firing retries automatically.
        var retryMsg = 'Transient error (retry ' + failCount + '/' + CRAWL_MAX_CONSEC_FAILURES + '): ' + e.message.slice(0, 60);
        if (es) setCrawlStatus(es, retryMsg);
      }
    } catch (e2) { /* swallow secondary error */ }
  } finally {
    lock.releaseLock();
  }
}

// Helper: handles the forward→backward (or forward→complete) transition.
// A crawl only ever runs ONE backward pass — without CRAWL_BACKWARD_DONE,
// this would re-enter 'backward' every time ANY forward phase completes,
// including the "expand backward papers" forward phase itself, oscillating
// forward<->backward indefinitely and pulling in ever-older, ever-more
// distant papers each cycle instead of stopping once genuinely done.
function transitionFromForward(props, sheet, runBackward, expandBackward, logRow) {
  var backwardAlreadyDone = props.getProperty('CRAWL_BACKWARD_DONE') === 'true';
  if (runBackward && !backwardAlreadyDone) {
    props.setProperty('CRAWL_PHASE',              'backward');
    props.setProperty('CRAWL_BACKWARD_IDX',       '0');
    props.setProperty('CRAWL_EXPANDING_BACKWARD', 'false');
    props.setProperty('CRAWL_BATCH_NUM',          '1');
    setCrawlStatus(sheet, 'Forward complete — starting backward pass…');
  } else {
    deleteCrawlTrigger();
    updateLogRow(logRow, 'Complete');
    setCrawlStatus(sheet, 'Complete');
  }
}

// ============================================================
// Sheet setup
// ============================================================

function newCrawlSheetName() {
  const now = new Date();
  const pad = function(n) { return n.toString().padStart(2, "0"); };
  return "Crawl " + now.getFullYear() + "-" + pad(now.getMonth() + 1) + "-" + pad(now.getDate()) +
         " " + pad(now.getHours()) + ":" + pad(now.getMinutes());
}

function setupCrawlSheet(sheet, direction, seedTitle) {
  sheet.clear().clearFormats();

  // Row 1: direction + seed label + status indicator
  sheet.getRange(1, 1).setValue(direction === "forward" ? "Forward crawl" : "Backward crawl")
    .setFontWeight("bold").setFontSize(11);
  sheet.getRange(1, 2).setValue(seedTitle || "").setFontColor("#888888").setFontStyle("italic");
  sheet.getRange(1, CRAWL_STATUS_LABEL_COL)
    .setValue("Crawl Status").setFontWeight("bold").setFontColor("#555555").setFontSize(10);
  // Value cell styled as "pending" until the first batch sets it
  sheet.getRange(1, CRAWL_STATUS_VALUE_COL)
    .setValue("Starting…").setBackground("#f4b400").setFontColor("#333333")
    .setFontWeight("bold").setFontSize(10);

  // Row 2 headers — three segments because cols 11 and 14 are MAP formulas, not static text.
  // Cols 1-10: static text (Depth … Cited By)
  sheet.getRange(2, 1, 1, CRAWL_COL.CITED_BY)
    .setValues([CRAWL_HEADERS.slice(0, CRAWL_COL.CITED_BY)])
    .setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  // Col 11: Filter Match MAP formula
  sheet.getRange(2, CRAWL_COL.FILTER_MATCH)
    .setFormula(CRAWL_DEFAULT_FILTER_FORMULA)
    .setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  // Cols 12-14: static text (Found From, Matched Cites, Direction)
  sheet.getRange(2, CRAWL_COL.FOUND_FROM, 1, 3)
    .setValues([CRAWL_HEADERS.slice(CRAWL_COL.FOUND_FROM - 1)])
    .setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  // Col 15: In-Sheet Links MAP formula
  sheet.getRange(2, CRAWL_IN_SHEET_LINKS_COL)
    .setFormula(CRAWL_IN_SHEET_LINKS_FORMULA)
    .setFontWeight("bold").setBackground("#4285f4").setFontColor("white");

  sheet.setFrozenRows(2);

  // Column widths
  sheet.setColumnWidth(CRAWL_COL.DEPTH,     60);
  sheet.setColumnWidth(CRAWL_COL.CRAWLED,   75);
  sheet.setColumnWidth(CRAWL_COL.YEAR,      60);
  sheet.setColumnWidth(CRAWL_COL.TITLE,     260);
  sheet.setColumnWidth(CRAWL_COL.AUTHORS,   180);
  sheet.setColumnWidth(CRAWL_COL.PUB_TYPES, 120);
  sheet.setColumnWidth(CRAWL_COL.VENUE,     130);
  sheet.setColumnWidth(CRAWL_COL.ABSTRACT,  400);
  sheet.setColumnWidth(CRAWL_COL.ID,        130);
  sheet.setColumnWidth(CRAWL_COL.CITED_BY,          80);
  sheet.setColumnWidth(CRAWL_COL.FILTER_MATCH,      110);
  sheet.setColumnWidth(CRAWL_COL.FOUND_FROM,         160);
  sheet.setColumnWidth(CRAWL_COL.MATCHED_CITES,      260);
  sheet.setColumnWidth(CRAWL_COL.DIRECTION,           55);
  sheet.setColumnWidth(CRAWL_IN_SHEET_LINKS_COL,     90);

  // Formatting for data rows
  const maxRows = sheet.getMaxRows();
  sheet.getRange(1, 1, maxRows, CRAWL_NUM_COLS).setWrap(true);
  sheet.getRange(1, CRAWL_COL.ABSTRACT, maxRows, 1)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.setRowHeights(3, maxRows - 2, CRAWL_ROW_HEIGHT);
  sheet.getRange(3, CRAWL_COL.TITLE,    maxRows - 2, 1).setFontSize(10);
  sheet.getRange(3, CRAWL_COL.AUTHORS,  maxRows - 2, 1).setFontSize(10);
  sheet.getRange(3, CRAWL_COL.ABSTRACT, maxRows - 2, 1).setFontSize(8);

  // Filter
  sheet.getRange(2, 1, maxRows - 1, CRAWL_NUM_COLS).createFilter();

  // CF rule: highlight entire row green when Filter Match = TRUE
  var cfRange = sheet.getRange(3, 1, maxRows - 2, CRAWL_NUM_COLS);
  var cfRule  = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$" + CRAWL_FILTER_MATCH_COL_LETTER + "3=TRUE")
    .setBackground("#b7e1cd")
    .setRanges([cfRange])
    .build();

  // CF rule: tint just the Abstract cell when it's blank — a missing
  // abstract means matching fell back to title text alone, so a "no match"
  // here is less certain than for a row with a full abstract. Formatting
  // only (no inserted text), so it can't itself be picked up by the term
  // search the way an inline marker in a searched column would be.
  var absLetter  = colToLetter(CRAWL_COL.ABSTRACT);
  var absRange   = sheet.getRange(3, CRAWL_COL.ABSTRACT, maxRows - 2, 1);
  var noAbsRule  = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$" + absLetter + "3=\"\"")
    .setBackground("#fff3cd")
    .setRanges([absRange])
    .build();

  sheet.setConditionalFormatRules([cfRule, noAbsRule]);
}

// ============================================================
// Sheet data helpers
// ============================================================

// Uses Depth column (A) to find last real data row, avoiding
// any formula-spill inflation of getLastRow().
function getCrawlLastDataRow(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < 3) return 2;
  const vals = sheet.getRange(3, CRAWL_COL.DEPTH, maxRows - 2, 1).getValues().flat();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (vals[i] !== "" && vals[i] !== null && vals[i] !== undefined) return i + 3;
  }
  return 2;
}

function getCrawlExistingIds(sheet) {
  const lastRow = getCrawlLastDataRow(sheet);
  if (lastRow < 3) return new Set();
  const ids = sheet.getRange(3, CRAWL_COL.ID, lastRow - 2, 1).getValues().flat();
  return new Set(ids.filter(Boolean));
}

// Returns the first row where Crawled = FALSE, or null if none.
function findNextUncrawled(sheet) {
  const lastRow = getCrawlLastDataRow(sheet);
  if (lastRow < 3) return null;
  const data = sheet.getRange(3, 1, lastRow - 2, CRAWL_NUM_COLS).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][CRAWL_COL.CRAWLED - 1] === false) {
      return {
        sheetRow: i + 3,
        depth:    data[i][CRAWL_COL.DEPTH - 1],
        id:       data[i][CRAWL_COL.ID - 1],
        title:    data[i][CRAWL_COL.TITLE - 1]
      };
    }
  }
  return null;
}

function countUncrawled(sheet) {
  const lastRow = getCrawlLastDataRow(sheet);
  if (lastRow < 3) return 0;
  const vals = sheet.getRange(3, CRAWL_COL.CRAWLED, lastRow - 2, 1).getValues().flat();
  return vals.filter(function(v) { return v === false; }).length;
}

// Returns the row number the batch was written at (or null if empty), so
// callers that need to attach anything post-write (e.g. abstract-source
// notes) know exactly which rows they ended up as, without recomputing —
// and possibly drifting from — the same position independently.
function writeCrawlRows(sheet, rows) {
  if (rows.length === 0) return null;
  const startRow = getCrawlLastDataRow(sheet) + 1;
  sheet.getRange(startRow, 1, rows.length, CRAWL_NUM_COLS).setValues(rows);
  sheet.getRange(startRow, CRAWL_COL.CRAWLED, rows.length, 1).insertCheckboxes();
  sheet.setRowHeights(startRow, rows.length, CRAWL_ROW_HEIGHT);
  var range = sheet.getRange(startRow, 1, rows.length, CRAWL_NUM_COLS);
  range.setWrap(true);
  sheet.getRange(startRow, CRAWL_COL.ABSTRACT, rows.length, 1)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(startRow, CRAWL_COL.TITLE,    rows.length, 1).setFontSize(10);
  sheet.getRange(startRow, CRAWL_COL.AUTHORS,  rows.length, 1).setFontSize(10);
  sheet.getRange(startRow, CRAWL_COL.ABSTRACT, rows.length, 1).setFontSize(8);
  return startRow;
}

// Attaches a hover note to the Abstract cell for rows where an
// abstract-source lookup was attempted — a cell note rather than a new
// column (which would shift every fixed CRAWL_COL position for any crawl
// sheet already in progress) or inline text (which would become part of
// the searched text now that title+abstract are matched together).
// notes[i] (aligned with the written rows, startRow+i) may be null/empty
// for rows where no lookup was needed (Semantic Scholar already had it).
function applyAbstractNotes(sheet, startRow, notes) {
  if (!startRow) return;
  for (var i = 0; i < notes.length; i++) {
    if (notes[i]) sheet.getRange(startRow + i, CRAWL_COL.ABSTRACT).setNote(notes[i]);
  }
}

// ============================================================
// Row builders
// ============================================================

// Flags likely parsing artifacts in backward-pass reference metadata (e.g.
// truncated citation fragments like ": The", "Beyond the", "Et al", or
// mixed-script noise) rather than genuine papers. Deliberately requires
// ALL of: no abstract, no authors, AND a title pattern typical of a
// truncated/malformed entry — a short title alone isn't enough, since
// real seminal papers can have very short titles ("Scratch",
// "Computational thinking", both genuine papers seen in this exact
// dataset) with no other signal that they're bogus.
function looksLikeMalformedReference(title, abstract, authors) {
  if (abstract || authors) return false;
  var t = (title || '').trim();
  if (!t) return true;
  // Starts with punctuation, e.g. ": The" — but quote marks are excluded,
  // since plenty of real titles legitimately open by quoting a phrase
  // (confirmed false positive: a real book review titled
  // "“Got TPACK?” If not, Here’s Where to Learn about It!").
  if (/^[^a-zA-Z0-9'"‘’“”«»]/.test(t)) return true;
  var words = t.toLowerCase().replace(/[.,;:]+$/, '').split(/\s+/);
  var lastWord = words[words.length - 1];
  var trailingStopwords = ['the', 'a', 'an', 'of', 'in', 'to', 'and', 'or', 'al', 'et'];
  if (trailingStopwords.indexOf(lastWord) !== -1) return true; // trailing fragment / "et al"
  if (/[一-鿿぀-ヿ가-힯]/.test(t)) return true; // CJK/Hangul mixed in
  return false;
}

function crawlRowFromS2(paper, depth, parentId, dir) {
  var mag = paper.externalIds && paper.externalIds.MAG;
  var id    = mag ? ("W" + mag) : ("S2:" + paper.paperId);
  var title = paper.title || "";
  if (looksLikeMalformedReference(title, paper.abstract, (paper.authors || []).length)) {
    title = '⚠ [unverified reference — check source] ' + title;
  }
  return [
    depth, false,
    paper.year || "",
    title,
    (paper.authors || []).map(function(a) { return a.name; }).join("; "),
    (paper.publicationTypes || []).join(", "),
    paper.venue || "",
    paper.abstract || "",
    id,
    paper.citationCount || 0,
    "",               // Filter Match — spilled from row 2 formula
    parentId || "",   // Found From
    "",               // Matched Cites — populated by updateCrawlMatchedCites
    dir || "F"        // Direction: F = forward pass, B = backward pass
  ];
}

function reconstructAbstract(invertedIndex) {
  if (!invertedIndex) return "";
  var words = [];
  var entries = Object.entries(invertedIndex);
  for (var i = 0; i < entries.length; i++) {
    var word      = entries[i][0];
    var positions = entries[i][1];
    for (var j = 0; j < positions.length; j++) {
      words[positions[j]] = word;
    }
  }
  return words.join(" ");
}

function crawlRowFromOpenAlex(work, depth, parentId, dir) {
  var id       = work.id.replace("https://openalex.org/", "");
  var abstract = reconstructAbstract(work.abstract_inverted_index);
  var venue    = (work.primary_location && work.primary_location.source)
                   ? work.primary_location.source.display_name : "";
  var title    = work.title || "";
  if (looksLikeMalformedReference(title, abstract, (work.authorships || []).length)) {
    title = '⚠ [unverified reference — check source] ' + title;
  }
  return [
    depth, false,
    work.publication_year || "",
    title,
    (work.authorships || []).map(function(a) { return a.author.display_name; }).join("; "),
    "",
    venue,
    abstract,
    id,
    work.cited_by_count || 0,
    "",               // Filter Match — spilled from row 2 formula
    parentId || "",   // Found From
    "",               // Matched Cites — populated by updateCrawlMatchedCites
    dir || "F"        // Direction
  ];
}

// ============================================================
// JS filter — mirrors the Sheets formula logic
// ============================================================

// Escapes regex metacharacters so a filter term can be dropped into a
// RegExp/REGEXMATCH pattern literally.
function escapeRegExpTerm(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function jsMatchesFilter(text, groups) {
  text = (text || "").toLowerCase();
  return groups.every(function(group) {
    if (!group.terms || !group.terms.trim()) return true;
    var terms = group.terms.split(",")
      .map(function(t) { return t.trim().replace(/^["']|["']$/g, "").trim().toLowerCase(); })
      .filter(function(t) { return t.length > 0; });
    if (terms.length === 0) return true;
    // Word-boundary match, not substring — a bare term like "AI" or "ML"
    // would otherwise match inside unrelated words ("said", "html").
    var anyMatch = terms.some(function(t) {
      return new RegExp('\\b' + escapeRegExpTerm(t) + '\\b', 'i').test(text);
    });
    return group.not ? !anyMatch : anyMatch;
  });
}

// ============================================================
// Fetch candidates
// ============================================================

// Forward: papers that cite the given paper (Semantic Scholar)
function fetchForwardCandidates(id, title) {
  return s2GetCitations(id, title);
}

// Back-off delays (ms) applied on transient OpenAlex failures (429 or 5xx).
// Shorter than S2's own [3000,10000,30000] since OpenAlex's polite pool
// (mailto set) is generally more permissive.
const OPENALEX_BACKOFF_MS = [1000, 3000, 8000];

// Best-effort fallback: recovers an abstract from OpenAlex when Semantic
// Scholar didn't have one, using whichever external ID is available (MAG,
// then DOI). Retries on rate-limit/server errors with back-off — confirmed
// cases where OpenAlex genuinely has the abstract but a single attempt came
// back empty, so this is worth persisting on rather than giving up at once.
// Gives up immediately on a non-retryable response (e.g. 404 — retrying
// won't produce a paper that doesn't exist under that ID).
//
// Returns { abstract, reason, url } rather than a bare string, so callers
// can record *why* a paper still has no abstract — distinguishing "never
// had an ID to check" from "checked, genuinely has none" from "the check
// itself failed and availability is unconfirmed" — for debugging which
// remaining gaps are real vs. a search failure worth retrying.
//   reason: 'no-id' | 'found' | 'no-abstract' | 'error'
//   url:    the paper's OpenAlex/DOI page (human-viewable), or '' for no-id
function fetchOpenAlexAbstract(externalIds) {
  if (!externalIds || (!externalIds.MAG && !externalIds.DOI)) {
    return { abstract: "", reason: "no-id", url: "" };
  }
  var url, pageUrl;
  if (externalIds.MAG) {
    pageUrl = "https://openalex.org/W" + externalIds.MAG;
    url     = "https://api.openalex.org/works/W" + externalIds.MAG +
              "?select=abstract_inverted_index&mailto=" + getOpenAlexEmail();
  } else {
    pageUrl = "https://doi.org/" + externalIds.DOI;
    url     = "https://api.openalex.org/works/https://doi.org/" +
              encodeURIComponent(externalIds.DOI) +
              "?select=abstract_inverted_index&mailto=" + getOpenAlexEmail();
  }

  var abstract = "";
  var reason   = "error";
  for (var attempt = 0; attempt <= OPENALEX_BACKOFF_MS.length; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var code = resp.getResponseCode();
      if (code === 200) {
        var data = JSON.parse(resp.getContentText());
        abstract = reconstructAbstract(data.abstract_inverted_index);
        reason   = abstract ? "found" : "no-abstract";
        break;
      }
      if ((code === 429 || code >= 500) && attempt < OPENALEX_BACKOFF_MS.length) {
        Utilities.sleep(OPENALEX_BACKOFF_MS[attempt]);
        continue;
      }
      break; // non-retryable (e.g. 404), or retries exhausted — reason stays 'error'
    } catch (e) {
      if (attempt < OPENALEX_BACKOFF_MS.length) {
        Utilities.sleep(OPENALEX_BACKOFF_MS[attempt]);
        continue;
      }
      break;
    }
  }
  // Brief pacing once settled (success or not) — a paper with many
  // missing-abstract references could otherwise fire several of these
  // back-to-back with no spacing at all, unlike the once-per-source-paper
  // sleep already covering the main S2 fetch.
  Utilities.sleep(300);
  return { abstract: abstract, reason: reason, url: pageUrl };
}

// Turns a fetchOpenAlexAbstract() result into a human-readable note for the
// Abstract cell — only called when Semantic Scholar itself had no abstract.
function describeAbstractSource(lookup) {
  switch (lookup.reason) {
    case "no-id":
      return "No abstract in Semantic Scholar, and no MAG/DOI available to check OpenAlex.";
    case "found":
      return "Abstract recovered from OpenAlex: " + lookup.url;
    case "no-abstract":
      return "No abstract in Semantic Scholar. Checked OpenAlex (" + lookup.url + ") — also has no abstract.";
    default: // 'error'
      return "No abstract in Semantic Scholar. OpenAlex check failed after retries (" +
             lookup.url + ") — availability unconfirmed, worth rechecking.";
  }
}

// Backward: papers that the given paper cites (OpenAlex referenced_works + batch metadata)
function fetchBackwardCandidates(openAlexId) {
  var url  = "https://api.openalex.org/works/" + openAlexId +
             "?mailto=" + getOpenAlexEmail() + "&select=id,referenced_works";
  var work = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
  var refs = (work.referenced_works || []).map(function(r) {
    return r.replace("https://openalex.org/", "");
  });
  if (refs.length === 0) return [];

  var works = [];
  for (var i = 0; i < refs.length; i += 50) {
    var batch     = refs.slice(i, i + 50);
    var batchUrl  = "https://api.openalex.org/works?filter=openalex_id:" +
                    batch.join("%7C") +
                    "&per_page=50&select=id,title,publication_year,authorships," +
                    "cited_by_count,abstract_inverted_index,primary_location" +
                    "&mailto=" + getOpenAlexEmail();
    var result    = JSON.parse(UrlFetchApp.fetch(batchUrl, { muteHttpExceptions: true }).getContentText());
    (result.results || []).forEach(function(w) { works.push(w); });
  }
  return works;
}

function getCandidateId(candidate, direction) {
  if (direction === "forward") {
    var mag = candidate.externalIds && candidate.externalIds.MAG;
    return mag ? ("W" + mag) : ("S2:" + candidate.paperId);
  }
  return candidate.id.replace("https://openalex.org/", "");
}

function getCandidateAbstract(candidate, direction) {
  if (direction === "forward") return candidate.abstract || "";
  return reconstructAbstract(candidate.abstract_inverted_index);
}

function getCandidateYear(candidate, direction) {
  return direction === "forward" ? candidate.year : candidate.publication_year;
}

// True when yearBound is off, the year is missing/unparseable (benefit of
// the doubt rather than penalising incomplete metadata), or the year falls
// within [yearFloor, yearCeiling]. yearFloor of 0 means "no seed had usable
// year metadata, and no explicit From year given" — no effective floor.
// yearCeiling of 0 means "no explicit To year given" — defaults to the
// current year rather than being unbounded.
function isYearInBounds(year, yearFloor, yearCeiling, yearBound) {
  if (!yearBound) return true;
  var y = parseInt(year);
  if (!y || isNaN(y)) return true;
  if (yearFloor && y < yearFloor) return false;
  var ceiling = yearCeiling || new Date().getFullYear();
  if (y > ceiling) return false;
  return true;
}

function buildCrawlRow(candidate, direction, depth, parentId, dir) {
  return direction === "forward"
    ? crawlRowFromS2(candidate, depth, parentId, dir)
    : crawlRowFromOpenAlex(candidate, depth, parentId, dir);
}

// Prefixes a paper's Title cell with a visible failure marker when its
// candidate fetch (citations or references) throws, so a genuinely-failed
// lookup is distinguishable from one that succeeded but found nothing new —
// both would otherwise leave the row Crawled=TRUE with zero new rows added,
// making them look identical in the exported data. Deliberately reuses the
// Title column rather than adding a new one, since a new column would shift
// every fixed column position (CRAWL_COL.*) and the term-helper/filter-formula
// columns that come after it, corrupting any crawl sheet created before this
// change if it's ever resumed.
function markFetchFailure(sheet, sheetRow, e) {
  var cell     = sheet.getRange(sheetRow, CRAWL_COL.TITLE);
  var curTitle = String(cell.getValue());
  var baseTitle = curTitle.replace(/^⚠ \[fetch failed:[^\]]*\]\s*/, '');
  cell.setValue('⚠ [fetch failed: ' + e.message.slice(0, 60) + '] ' + baseTitle);
}

// ============================================================
// Main crawl loop
// ============================================================

// paperDir: direction marker written into the Direction column for newly-added rows.
// "F" for forward-pass papers (default), "B" when expanding backward-discovered papers.
// matchesOnly: when false, non-matching candidates are also written (as
// already-Crawled rows, so they aren't re-queued) instead of being discarded.
// yearFloor/yearBound: when yearBound is on, candidates older than yearFloor
// (or newer than the current year) are treated as non-matches — logged as a
// dead end (matchesOnly's existing Crawled=TRUE handling) rather than
// excluded outright, but never expanded further.
function runCrawlLoop(sheet, direction, groups, maxDepth, maxPapers, paperDir, matchesOnly, yearFloor, yearCeiling, yearBound) {
  paperDir    = paperDir || "F";
  matchesOnly = matchesOnly !== false;
  var startTime = Date.now();
  var processed = 0;

  while (true) {
    // Time guard — stop before Apps Script kills the process
    if (Date.now() - startTime > CRAWL_TIME_LIMIT_MS) {
      updateCrawlMatchedCites(sheet);
      var remaining = countUncrawled(sheet);
      return { status: 'time-limit', message: "Time limit reached. Processed " + processed + " papers this session; " +
             remaining + " remain in queue. Click Resume Crawl to continue." };
    }

    var next = findNextUncrawled(sheet);
    if (!next) {
      updateCrawlMatchedCites(sheet);
      return { status: 'complete', message: "Crawl complete. " + processed + " papers processed in this session." };
    }

    var sheetRow = next.sheetRow;
    var depth    = next.depth;
    var id       = next.id;
    var title    = next.title;

    // At max depth: mark done without fetching
    if (depth >= maxDepth) {
      sheet.getRange(sheetRow, CRAWL_COL.CRAWLED).setValue(true);
      processed++;
      continue;
    }

    // Fetch candidates, skip paper on error
    var candidates = [];
    try {
      candidates = direction === "forward"
        ? fetchForwardCandidates(id, title)
        : fetchBackwardCandidates(id);
    } catch (e) {
      markFetchFailure(sheet, sheetRow, e);
      sheet.getRange(sheetRow, CRAWL_COL.CRAWLED).setValue(true);
      processed++;
      continue;
    }

    // Pace S2 requests — stay within rate limits even without an API key
    if (direction === 'forward') Utilities.sleep(1100);

    // Filter: not already in sheet; matches are queued for further expansion,
    // non-matches (when matchesOnly is off) are recorded but marked Crawled
    // so they aren't independently expanded.
    var existingIds   = getCrawlExistingIds(sheet);
    var matchRows     = [];
    var abstractNotes = []; // aligned with matchRows — note text or null per row
    for (var i = 0; i < candidates.length; i++) {
      var c   = candidates[i];
      var cId = getCandidateId(c, direction);
      if (existingIds.has(cId)) continue;
      var abstract = getCandidateAbstract(c, direction);
      var note     = null;

      // Try to recover any missing abstract from OpenAlex, regardless of
      // whether it would change the match outcome — the goal is minimising
      // gaps in the sheet's own data, not just correcting matches.
      // Forward-only: the OpenAlex candidate path is unused by the current UI.
      if (!abstract && direction === "forward") {
        var lookup = fetchOpenAlexAbstract(c.externalIds);
        if (lookup.abstract) {
          abstract   = lookup.abstract;
          c.abstract = lookup.abstract; // so the written row reflects the recovered text
        }
        note = describeAbstractSource(lookup);
      }

      var yearOk  = isYearInBounds(getCandidateYear(c, direction), yearFloor, yearCeiling, yearBound);
      var isMatch = jsMatchesFilter((c.title || '') + ' ' + abstract, groups) && yearOk;
      if (!isMatch && matchesOnly) continue;
      var row = buildCrawlRow(c, direction, depth + 1, id, paperDir);
      if (!isMatch) row[CRAWL_COL.CRAWLED - 1] = true;
      matchRows.push(row);
      abstractNotes.push(note);
      existingIds.add(cId); // deduplicate within this batch before writing
    }

    // Respect the paper cap
    if (matchRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd       = Math.max(0, maxPapers - currentCount);
      var writeStart   = writeCrawlRows(sheet, matchRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, abstractNotes.slice(0, canAdd));
      if (canAdd < matchRows.length) {
        sheet.getRange(sheetRow, CRAWL_COL.CRAWLED).setValue(true);
        processed++;
        updateCrawlMatchedCites(sheet);
        var logRowPL = parseInt(PropertiesService.getScriptProperties().getProperty('CRAWL_LOG_ROW') || '0') || 0;
        updateLogRow(logRowPL, 'Paper Limit');
        return { status: 'paper-limit', message: "Paper limit (" + maxPapers + ") reached. The queue still has unprocessed papers — " +
               "click Resume Crawl to continue (existing queue only; no new papers will be added)." };
      }
    }

    sheet.getRange(sheetRow, CRAWL_COL.CRAWLED).setValue(true);
    processed++;
  }
}

// ============================================================
// Seed resolution — pure S2, flexible input formats
// ============================================================

// Accepts DOI, OpenAlex W-ID, ArXiv ID, bare S2 paper ID, or any URL
// form of the above.  Returns a full S2 paper object or null if the
// input format is not recognised.  Throws on HTTP / API errors.
function resolveToS2Seed(input) {
  var clean = input.trim();
  var s2Ref; // identifier string passed to the S2 /paper/ endpoint

  // S2 paper URL  e.g. https://www.semanticscholar.org/paper/Title.../abc123...
  var s2UrlMatch = clean.match(/semanticscholar\.org\/paper\/[^/]+\/([a-f0-9]{40})\b/i);
  if (s2UrlMatch) {
    s2Ref = s2UrlMatch[1];
  }
  // Bare S2 paper ID (40 hex chars)
  else if (/^[a-f0-9]{40}$/i.test(clean)) {
    s2Ref = clean;
  }
  // Our own "S2:<hash>" sheet/log ID format (see getCandidateId / startCrawl)
  else if (/^S2:[a-f0-9]{40}$/i.test(clean)) {
    s2Ref = clean.replace(/^S2:/i, '');
  }
  // ArXiv URL  e.g. https://arxiv.org/abs/2301.00001
  else if (/arxiv\.org\/abs\//i.test(clean)) {
    s2Ref = 'ArXiv:' + clean.replace(/.*arxiv\.org\/abs\//i, '').replace(/v\d+$/, '').trim();
  }
  // ArXiv ID — new format  e.g. 2301.00001 or arXiv:2301.00001v2
  else if (/^(arxiv:)?\d{4}\.\d{4,5}(v\d+)?$/i.test(clean)) {
    s2Ref = 'ArXiv:' + clean.replace(/^arxiv:/i, '').replace(/v\d+$/, '');
  }
  // DOI URL  e.g. https://doi.org/10.1000/xyz
  else if (/doi\.org\//i.test(clean)) {
    s2Ref = 'DOI:' + clean.replace(/.*doi\.org\//i, '');
  }
  // Bare DOI  e.g. 10.1000/xyz
  else if (/^10\.\d{4,}\//.test(clean)) {
    s2Ref = 'DOI:' + clean;
  }
  // OpenAlex URL  e.g. https://openalex.org/W2741809807
  else if (/openalex\.org\/W(\d+)/i.test(clean)) {
    s2Ref = 'MAG:' + clean.match(/W(\d+)/i)[1];
  }
  // Bare OpenAlex W-ID  e.g. W2741809807
  else if (/^W\d+$/i.test(clean)) {
    s2Ref = 'MAG:' + clean.slice(1);
  }
  else {
    return null; // unrecognised format
  }

  var fields = 'paperId,externalIds,title,abstract,year,authors,citationCount,publicationTypes,venue';
  var url    = 'https://api.semanticscholar.org/graph/v1/paper/' +
               encodeURIComponent(s2Ref) + '?fields=' + fields;
  var resp   = s2Fetch(url);
  var code   = resp.getResponseCode();
  var text   = resp.getContentText();
  if (code !== 200) {
    throw new Error('Semantic Scholar returned HTTP ' + code + ' for: ' + clean);
  }
  var paper = JSON.parse(text);
  if (paper.error || paper.message) throw new Error(paper.error || paper.message);
  if (!paper.paperId) return null;
  return paper;
}

// ============================================================
// Seed paper lookup — step-by-step, called from "Find Paper" button
// ============================================================

// Tries to find the paper on S2 and returns { steps, paper } or { steps, error }.
// steps = [{ text, ok }] where ok is true/false/null (null = informational).
// Falls back to S2 title search when no ID format is recognised.
function findSeedPaper(input) {
  var clean = (input || '').trim();
  if (!clean) return { steps: [], error: 'Please enter a search term.' };

  var steps  = [];
  var s2Ref  = null;

  // --- Format detection (same logic as resolveToS2Seed) ---
  var s2UrlMatch = clean.match(/semanticscholar\.org\/paper\/[^/]+\/([a-f0-9]{40})\b/i);
  if (s2UrlMatch) {
    s2Ref = s2UrlMatch[1];
    steps.push({ text: 'Recognised as Semantic Scholar URL', ok: true });
  } else if (/^[a-f0-9]{40}$/i.test(clean)) {
    s2Ref = clean;
    steps.push({ text: 'Recognised as Semantic Scholar paper ID', ok: true });
  } else if (/^S2:[a-f0-9]{40}$/i.test(clean)) {
    s2Ref = clean.replace(/^S2:/i, '');
    steps.push({ text: 'Recognised as Semantic Scholar paper ID', ok: true });
  } else if (/arxiv\.org\/abs\//i.test(clean)) {
    s2Ref = 'ArXiv:' + clean.replace(/.*arxiv\.org\/abs\//i, '').replace(/v\d+$/, '').trim();
    steps.push({ text: 'Recognised as ArXiv URL → ' + s2Ref, ok: true });
  } else if (/^(arxiv:)?\d{4}\.\d{4,5}(v\d+)?$/i.test(clean)) {
    s2Ref = 'ArXiv:' + clean.replace(/^arxiv:/i, '').replace(/v\d+$/, '');
    steps.push({ text: 'Recognised as ArXiv ID → ' + s2Ref, ok: true });
  } else if (/doi\.org\//i.test(clean)) {
    s2Ref = 'DOI:' + clean.replace(/.*doi\.org\//i, '');
    steps.push({ text: 'Recognised as DOI URL → ' + s2Ref, ok: true });
  } else if (/^10\.\d{4,}\//.test(clean)) {
    s2Ref = 'DOI:' + clean;
    steps.push({ text: 'Recognised as DOI → ' + s2Ref, ok: true });
  } else if (/openalex\.org\/W(\d+)/i.test(clean)) {
    s2Ref = 'MAG:' + clean.match(/W(\d+)/i)[1];
    steps.push({ text: 'Recognised as OpenAlex URL → ' + s2Ref, ok: true });
  } else if (/^W\d+$/i.test(clean)) {
    s2Ref = 'MAG:' + clean.slice(1);
    steps.push({ text: 'Recognised as OpenAlex W-ID → ' + s2Ref, ok: true });
  }

  // --- Title search fallback ---
  if (!s2Ref) {
    steps.push({ text: 'No ID format recognised — searching by title on Semantic Scholar', ok: null });
    var sf  = 'paperId,externalIds,title,abstract,year,authors,citationCount,publicationTypes,venue';
    var su  = 'https://api.semanticscholar.org/graph/v1/paper/search?query=' +
              encodeURIComponent(clean) + '&fields=' + sf + '&limit=1';
    var sr  = UrlFetchApp.fetch(su, getS2FetchOptions());
    var sc  = sr.getResponseCode();
    if (sc === 200) {
      var sd  = JSON.parse(sr.getContentText());
      var hit = sd.data && sd.data[0];
      if (hit && hit.paperId) {
        steps.push({ text: 'Found via title search: "' + (hit.title || hit.paperId) + '"', ok: true });
        return { steps: steps, paper: hit };
      }
      steps.push({ text: 'No results found — try a DOI or ID for precision', ok: false });
      return { steps: steps, error: 'No paper found. Try a more specific title or paste a DOI / ID.' };
    }
    steps.push({ text: 'Title search failed (HTTP ' + sc + ')', ok: false });
    return { steps: steps, error: 'Search failed (HTTP ' + sc + '). Please try again.' };
  }

  // --- Direct lookup with exponential back-off (3 s → 10 s → 30 s) ---
  var fields = 'paperId,externalIds,title,abstract,year,authors,citationCount,publicationTypes,venue';
  var url    = 'https://api.semanticscholar.org/graph/v1/paper/' +
               encodeURIComponent(s2Ref) + '?fields=' + fields;
  steps.push({ text: 'Looking up ' + s2Ref + ' on Semantic Scholar', ok: null });

  var waitLabels = ['3 s', '10 s', '30 s'];
  for (var attempt = 0; attempt <= S2_BACKOFF_MS.length; attempt++) {
    var resp = UrlFetchApp.fetch(url, getS2FetchOptions());
    var code = resp.getResponseCode();
    if (code === 200) {
      var paper = JSON.parse(resp.getContentText());
      if (paper.error || paper.message) {
        steps.push({ text: 'API error: ' + (paper.error || paper.message), ok: false });
        return { steps: steps, error: paper.error || paper.message };
      }
      steps.push({ text: 'Found: "' + paper.title + '"', ok: true });
      return { steps: steps, paper: paper };
    }
    if (code === 429 && attempt < S2_BACKOFF_MS.length) {
      steps.push({ text: 'Rate limited — waiting ' + waitLabels[attempt] + ' then retrying', ok: null });
      Utilities.sleep(S2_BACKOFF_MS[attempt]);
      continue;
    }
    var msg = code === 429
      ? 'Still rate limited after ' + (attempt) + ' retries — please wait a minute and try again'
      : 'Not found (HTTP ' + code + ') — check the ID or try a different format';
    steps.push({ text: msg, ok: false });
    return {
      steps: steps,
      error: code === 429
        ? 'Semantic Scholar is still rate limiting. Wait ~1 minute before trying again.'
        : 'Paper not found (HTTP ' + code + ').'
    };
  }

  return { steps: steps, error: 'Unexpected error.' };
}

// ============================================================
// S2 references fetch (backward pass)
// ============================================================

// Fetches the reference list (papers cited BY paperSheetId) from Semantic Scholar.
// paperSheetId is in sheet format: "W{mag}" or "S2:{hex}".
function s2GetReferences(paperSheetId) {
  var s2Id;
  if (/^W\d+$/.test(paperSheetId)) {
    s2Id = 'MAG:' + paperSheetId.slice(1);
  } else if (paperSheetId.indexOf('S2:') === 0) {
    s2Id = paperSheetId.slice(3);
  } else {
    return []; // unrecognised format
  }

  var fields  = 'paperId,externalIds,title,abstract,year,authors,citationCount,publicationTypes,venue';
  var limit   = 500;
  var offset  = 0;
  var papers  = [];

  while (true) {
    if (offset > 0) Utilities.sleep(1100);
    var url  = 'https://api.semanticscholar.org/graph/v1/paper/' +
               encodeURIComponent(s2Id) + '/references?fields=' + fields +
               '&limit=' + limit + '&offset=' + offset;
    var resp = s2Fetch(url);
    var code = resp.getResponseCode();
    if (code !== 200) throw new Error('S2 /references HTTP ' + code + ' for ' + paperSheetId);
    var data = JSON.parse(resp.getContentText());
    if (data.error || data.message) throw new Error(data.error || data.message);
    (data.data || []).forEach(function(item) {
      if (item.citedPaper && item.citedPaper.paperId) papers.push(item.citedPaper);
    });
    if ((data.data || []).length < limit) break;
    offset += limit;
  }

  return papers;
}

// ============================================================
// Backward pass
// ============================================================

// For each paper already in the sheet (all directions), fetches its references
// via S2 and adds new papers that pass the filter with Direction="B".
// When expandBackward=true those papers get Crawled=FALSE so the forward loop
// will later fetch their citations.  Batches against CRAWL_TIME_LIMIT_MS and
// persists progress in CRAWL_BACKWARD_IDX between trigger invocations.
//
// backwardDepth bounds this independently of (and typically much smaller
// than) the forward maxDepth. Without it, every new row this pass writes
// gets appended to the same sheet it re-scans every batch, so anything
// below the forward maxDepth stays eligible to have ITS references looked
// up too — turning "one supplementary backward hop" into a second
// full-depth crawl that recursively snowballs for as long as it keeps
// finding new candidates.
function runBackwardPass(sheet, groups, backwardDepth, maxPapers, expandBackward, matchesOnly, yearFloor, yearCeiling, yearBound) {
  matchesOnly = matchesOnly !== false;
  var startTime = Date.now();
  var props     = PropertiesService.getScriptProperties();
  var idx       = parseInt(props.getProperty('CRAWL_BACKWARD_IDX') || '0');

  var lastRow = getCrawlLastDataRow(sheet);
  if (lastRow < 3) return { status: 'complete', message: 'Backward pass complete. No papers found.' };

  var numRows = lastRow - 2;
  var data    = sheet.getRange(3, 1, numRows, CRAWL_NUM_COLS).getValues();

  // Build current ID set for dedup and collect all paper IDs to process
  // (paperSheetRows tracks each ID's real row — rows with a blank ID are
  // skipped, so the index into paperIds doesn't line up with row number).
  // paperDepths tracks each ID's own Depth, so newly-discovered references
  // inherit "parent depth + 1" instead of always resetting to 0 — otherwise
  // backwardDepth (below) stops bounding anything once backward discovery
  // kicks in.
  var allIds         = new Set();
  var paperIds        = [];
  var paperSheetRows  = [];
  var paperDepths     = [];
  data.forEach(function(row, i) {
    var id = String(row[CRAWL_COL.ID - 1] || '').trim();
    if (id) {
      allIds.add(id);
      paperIds.push(id);
      paperSheetRows.push(3 + i);
      paperDepths.push(parseInt(row[CRAWL_COL.DEPTH - 1]) || 0);
    }
  });

  var processed = 0;

  while (idx < paperIds.length) {
    if (Date.now() - startTime > CRAWL_TIME_LIMIT_MS) {
      props.setProperty('CRAWL_BACKWARD_IDX', String(idx));
      updateCrawlMatchedCites(sheet);
      return { status: 'time-limit', message: 'Time limit reached in backward pass. Processed ' + processed +
             ' papers; ' + (paperIds.length - idx) + ' remain.' };
    }

    var paperSheetRow  = paperSheetRows[idx];
    var paperDepth     = paperDepths[idx];
    var paperId        = paperIds[idx++];
    processed++;

    // Once a paper's own depth reaches backwardDepth, don't explore further
    // from it — this is the cap that keeps the backward pass to a bounded
    // number of hops instead of recursively following every paper it finds.
    if (paperDepth >= backwardDepth) continue;

    Utilities.sleep(1100);

    var refs = [];
    try { refs = s2GetReferences(paperId); } catch (e) {
      markFetchFailure(sheet, paperSheetRow, e);
      continue;
    }

    var newRows      = [];
    var newRowNotes  = []; // aligned with newRows — note text or null per row
    refs.forEach(function(ref) {
      if (!ref || !ref.paperId) return;
      var mag   = ref.externalIds && ref.externalIds.MAG;
      var refId = mag ? ('W' + mag) : ('S2:' + ref.paperId);
      if (allIds.has(refId)) return;
      var note = null;

      // Same fallback as the forward loop — try to recover any missing
      // abstract from OpenAlex regardless of whether it would change the
      // match outcome, to minimise gaps in the sheet's own data.
      if (!ref.abstract) {
        var lookup = fetchOpenAlexAbstract(ref.externalIds);
        if (lookup.abstract) ref.abstract = lookup.abstract;
        note = describeAbstractSource(lookup);
      }

      var yearOk  = isYearInBounds(ref.year, yearFloor, yearCeiling, yearBound);
      var isMatch = jsMatchesFilter((ref.title || '') + ' ' + (ref.abstract || ''), groups) && yearOk;
      if (!isMatch && matchesOnly) return;
      allIds.add(refId);
      var row = crawlRowFromS2(ref, paperDepth + 1, paperId, 'B');
      // Mark already-crawled (skip re-queueing) if not expanding, or if this
      // is a non-match recorded only for visibility.
      if (!expandBackward || !isMatch) row[CRAWL_COL.CRAWLED - 1] = true;
      newRows.push(row);
      newRowNotes.push(note);
    });

    if (newRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd       = Math.max(0, maxPapers - currentCount);
      var writeStart   = writeCrawlRows(sheet, newRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, newRowNotes.slice(0, canAdd));
    }
  }

  props.setProperty('CRAWL_BACKWARD_IDX', String(idx));
  updateCrawlMatchedCites(sheet);
  return { status: 'complete', message: 'Backward pass complete. Processed ' + processed + ' papers.' };
}

// ============================================================
// Public entry points (called from panel via google.script.run)
// ============================================================

// seeds   = array of S2 paper objects (already resolved by the panel via findSeedPaper).
// options = { runBackward: bool, expandBackward: bool }
// All seeds are written as depth-0 Direction="F" rows; the trigger loop handles phases.
function startCrawl(seeds, direction, maxDepth, maxPapers, groups, crawlName, options) {
  try {
    if (!seeds || !seeds.length) return 'No seed papers provided.';
    var opts           = options || {};
    var runBackward    = !!opts.runBackward;
    var expandBackward = !!opts.expandBackward;
    var matchesOnly    = opts.matchesOnly !== false;
    var yearBound      = opts.yearBound   !== false;
    // Independent from (and typically much smaller than) the forward maxDepth
    // above — bounds how many hops of reference-lookups the backward pass
    // itself recursively follows, so it can't turn into a second full-depth
    // crawl (runBackwardPass rebuilds its candidate queue from the whole
    // sheet every batch, so anything it finds is otherwise eligible to have
    // its own references looked up too).
    var backwardDepth  = parseInt(opts.backwardDepth) || 1;

    // Floor defaults to the earliest seed year, ceiling to the present day —
    // bounding the corpus so backward-discovered references don't wander
    // arbitrarily far into the past. Both can be overridden with an explicit
    // From/To year from the panel; 0 means "no override, use the default".
    var seedYears = seeds
      .map(function(s) { return parseInt(s.year); })
      .filter(function(y) { return y && !isNaN(y); });
    var autoYearFloor = seedYears.length ? Math.min.apply(null, seedYears) : 0;
    var yearFromOverride = parseInt(opts.yearFrom) || 0;
    var yearToOverride   = parseInt(opts.yearTo)   || 0;
    var yearFloor   = yearFromOverride || autoYearFloor;
    var yearCeiling = yearToOverride; // 0 = no override, isYearInBounds defaults to present day

    var sheetName = (crawlName || '').trim() || newCrawlSheetName();
    var seedLabel = seeds.length === 1
      ? (seeds[0].title || 'Unknown')
      : seeds.length + ' seed papers';
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.insertSheet(sheetName);
    setupCrawlSheet(sheet, direction, seedLabel);

    var props = PropertiesService.getScriptProperties();
    props.setProperty('CRAWL_ACTIVE_SHEET',     sheetName);
    props.setProperty('CRAWL_DIRECTION',        direction);
    props.setProperty('CRAWL_MAX_DEPTH',        String(maxDepth));
    props.setProperty('CRAWL_MAX_PAPERS',       String(maxPapers));
    props.setProperty('SNOWBALL_FILTER_GROUPS', JSON.stringify(groups));
    props.setProperty('CRAWL_PHASE',            'forward');
    props.setProperty('CRAWL_RUN_BACKWARD',     runBackward    ? 'true' : 'false');
    props.setProperty('CRAWL_EXPAND_BACKWARD',  expandBackward ? 'true' : 'false');
    props.setProperty('CRAWL_MATCHES_ONLY',     matchesOnly    ? 'true' : 'false');
    props.setProperty('CRAWL_YEAR_BOUND',       yearBound      ? 'true' : 'false');
    props.setProperty('CRAWL_YEAR_FLOOR',       String(yearFloor));
    props.setProperty('CRAWL_YEAR_CEILING',     String(yearCeiling));
    props.setProperty('CRAWL_BACKWARD_DEPTH',   String(backwardDepth));
    // These are script-wide properties, not scoped to a single crawl sheet —
    // without resetting them here, a fresh crawl can inherit 'true' left
    // over from a previous crawl's backward-expansion phase, mislabeling
    // this crawl's own forward-citation rows as Direction='B' from the start.
    props.setProperty('CRAWL_EXPANDING_BACKWARD', 'false');
    props.setProperty('CRAWL_BACKWARD_IDX',       '0');
    // Also script-wide — without resetting, a fresh crawl could inherit
    // 'true' from an earlier crawl and skip its own (first) backward pass.
    props.setProperty('CRAWL_BACKWARD_DONE',       'false');

    // Log the crawl — store the row number so the trigger can update status later
    var seedIds = seeds.map(function(seed) {
      var mag = seed.externalIds && seed.externalIds.MAG;
      return mag ? ('W' + mag) : ('S2:' + seed.paperId);
    });
    var logRow = appendLogRow('Crawl', {
      name:          sheetName,
      seeds:         seedIds,
      depth:         maxDepth,
      maxPapers:     maxPapers,
      filterGroups:  groups,
      runBackward:   runBackward,
      expandBackward:expandBackward
    });
    if (logRow) props.setProperty('CRAWL_LOG_ROW', String(logRow));

    // Build and write all seed rows in one batch (Direction = "F")
    var seedRows = seeds.map(function(seed) {
      var mag    = seed.externalIds && seed.externalIds.MAG;
      var seedId = mag ? ('W' + mag) : ('S2:' + seed.paperId);
      return [
        0, false,
        seed.year  || '',
        seed.title || '',
        (seed.authors          || []).map(function(a) { return a.name; }).join('; '),
        (seed.publicationTypes || []).join(', '),
        seed.venue    || '',
        seed.abstract || '',
        seedId,
        seed.citationCount || 0,
        '', '', '', 'F'  // Filter Match (spill), Found From, Matched Cites, Direction
      ];
    });

    sheet.getRange(3, 1, seedRows.length, CRAWL_NUM_COLS).setValues(seedRows);
    sheet.getRange(3, CRAWL_COL.CRAWLED, seedRows.length, 1).insertCheckboxes();
    sheet.setRowHeights(3, seedRows.length, CRAWL_ROW_HEIGHT);

    ss.setActiveSheet(sheet);
    applyCrawlHighlight(sheet, groups);

    // Start the background trigger loop
    props.setProperty('CRAWL_BATCH_NUM', '1');
    createCrawlTrigger();
    setCrawlStatus(sheet, 'Running forward batch 1…');

    return 'Crawl started — running in the background. ' +
           'Watch the "Crawl Status" cell at the top of the sheet. ' +
           'You can close this panel.';

  } catch (e) {
    return 'Error: ' + e.message;
  }
}

function resumeCrawl() {
  try {
    var props     = PropertiesService.getScriptProperties();
    var sheetName = props.getProperty("CRAWL_ACTIVE_SHEET");
    if (!sheetName) return "No active crawl found. Start a new crawl first.";

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return "Crawl sheet \"" + sheetName + "\" not found — it may have been deleted.";

    var direction = props.getProperty("CRAWL_DIRECTION")       || "forward";
    var groups    = JSON.parse(props.getProperty("SNOWBALL_FILTER_GROUPS") || "[]");
    var maxDepth  = parseInt(props.getProperty("CRAWL_MAX_DEPTH")  || "2");
    var maxPapers = parseInt(props.getProperty("CRAWL_MAX_PAPERS") || "300");

    var remaining = countUncrawled(sheet);
    if (remaining === 0) return "Crawl \"" + sheetName + "\" is already complete.";

    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);

    // Resume from wherever the batch counter left off
    var batch = parseInt(props.getProperty('CRAWL_BATCH_NUM') || '1');
    createCrawlTrigger();
    setCrawlStatus(sheet, 'Running batch ' + batch + '…');

    return 'Crawl resumed — running in the background. ' +
           'Watch the "Crawl Status" cell at the top of the sheet.';

  } catch (e) {
    return "Error: " + e.message;
  }
}

// ============================================================
// Relationship tracking
// ============================================================

// Reads the Found From column and writes comma-separated child IDs back into
// the Matched Cites column for each parent paper.  Called at the end of every
// crawl session (time-limit, complete, or paper-cap) so the column stays current.
function updateCrawlMatchedCites(sheet) {
  var lastRow = getCrawlLastDataRow(sheet);
  if (lastRow < 3) return;
  var numRows = lastRow - 2;

  // Read ID (col 9) and Found From (col 12) in one batch
  var idCol = sheet.getRange(3, CRAWL_COL.ID,         numRows, 1).getValues();
  var ffCol = sheet.getRange(3, CRAWL_COL.FOUND_FROM, numRows, 1).getValues();

  // Build reverse map: parentId → [child IDs]
  var parentToChildren = {};
  for (var i = 0; i < numRows; i++) {
    var childId  = String(idCol[i][0] || "").trim();
    var parentId = String(ffCol[i][0] || "").trim();
    if (!parentId || !childId) continue;
    if (!parentToChildren[parentId]) parentToChildren[parentId] = [];
    parentToChildren[parentId].push(childId);
  }

  // Write Matched Cites for every row
  var updates = [];
  for (var j = 0; j < numRows; j++) {
    var id       = String(idCol[j][0] || "").trim();
    var children = parentToChildren[id] || [];
    updates.push([children.join(", ")]);
  }
  if (updates.length > 0) {
    sheet.getRange(3, CRAWL_COL.MATCHED_CITES, updates.length, 1).setValues(updates);
  }
}

// ============================================================
// Conditional formatting helpers
// ============================================================

// Writes per-term helper columns + Filter Match formula onto the crawl sheet.
// Abstract = col 8 (H); Filter Match = col 11 (K); term columns start at col 12 (L).
// Reuses colToLetter / buildTermFormula / buildInnerExpression from snowball.js.
function applyCrawlHighlight(sheet, groups) {
  var CRAWL_FIRST_DETAIL = CRAWL_FIRST_DETAIL_COL;  // col 15 = O
  var absLetter = colToLetter(CRAWL_COL.ABSTRACT);   // H

  // Always keep the In-Sheet Links formula current (covers re-apply after new crawl data)
  sheet.getRange(2, CRAWL_IN_SHEET_LINKS_COL)
    .setFormula(CRAWL_IN_SHEET_LINKS_FORMULA)
    .setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setColumnWidth(CRAWL_IN_SHEET_LINKS_COL, 90);

  // Parse terms, drop empty groups
  var parsed = groups.map(function(g) {
    return {
      not:   g.not,
      terms: (g.terms || '').split(',')
        .map(function(t) { return t.trim().replace(/^["'"']|["'"']$/g, '').trim(); })
        .filter(function(t) { return t.length > 0; })
    };
  }).filter(function(g) { return g.terms.length > 0; });

  if (parsed.length === 0) return;

  // Clear previous helper area first
  var maxHelper = 60;
  sheet.getRange(1, CRAWL_FIRST_DETAIL, 1, maxHelper).breakApart().clearContent().clearFormat();
  sheet.getRange(2, CRAWL_FIRST_DETAIL, 1, maxHelper).clearContent().clearFormat();

  var colNum = CRAWL_FIRST_DETAIL;
  parsed.forEach(function(g, groupIdx) {
    var groupStartCol = colNum;
    var isNot  = g.not;
    var termBg = isNot ? '#fce8e6' : '#e8f0fe';

    g.terms.forEach(function(term) {
      sheet.getRange(2, colNum)
        .setFormula(buildTermFormula(term, CRAWL_COL.TITLE, CRAWL_COL.ABSTRACT))
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('bottom')
        .setTextRotation(90)
        .setBackground(termBg)
        .setFontColor('#222');

      sheet.setColumnWidth(colNum, 35);
      colNum++;
    });

    // Merged group label in row 1
    if (g.terms.length > 0) {
      var label  = (isNot ? 'NOT Group ' : 'Group ') + (groupIdx + 1);
      var hdrBg  = isNot ? '#e53935' : '#1a73e8';
      sheet.getRange(1, groupStartCol, 1, g.terms.length)
        .merge()
        .setValue(label)
        .setBackground(hdrBg)
        .setFontColor('white')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

      sheet.getRange(1, groupStartCol, sheet.getMaxRows(), 1)
        .setBorder(null, true, null, null, null, null,
                   '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  });

  if (colNum > CRAWL_FIRST_DETAIL) {
    sheet.setRowHeight(2, 130);

    // CF: TRUE → yellow block, FALSE → white on white (invisible)
    var termDataRange = sheet.getRange(3, CRAWL_FIRST_DETAIL, sheet.getMaxRows() - 2, colNum - CRAWL_FIRST_DETAIL);
    var startLetter = colToLetter(CRAWL_FIRST_DETAIL);
    var trueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + startLetter + '3=TRUE')
      .setBackground('#FFF176').setFontColor('#FFF176')
      .setRanges([termDataRange]).build();
    var falseRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + startLetter + '3=FALSE')
      .setBackground('#FFFFFF').setFontColor('#FFFFFF')
      .setRanges([termDataRange]).build();
    // Keep the green row-highlight rule; drop any stale helper-column rules.
    var existingRules = sheet.getConditionalFormatRules().filter(function(rule) {
      return !rule.getRanges().some(function(r) { return r.getColumn() >= CRAWL_FIRST_DETAIL; });
    });
    sheet.setConditionalFormatRules(existingRules.concat([trueRule, falseRule]));
  }

  // Filter Match: references per-term columns — no duplicated calculation
  var filterFormula = buildFilterMatchFromTermCols(parsed, CRAWL_FIRST_DETAIL);
  if (!filterFormula) return;
  sheet.getRange(2, CRAWL_COL.FILTER_MATCH)
    .setFormula(filterFormula)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');
}

// Called from the crawl panel "Apply Highlight Rule" button.
// Updates the filter formula and CF rule on the active crawl sheet.
function applyCrawlFilter(groups) {
  try {
    var props     = PropertiesService.getScriptProperties();
    var sheetName = props.getProperty("CRAWL_ACTIVE_SHEET");
    if (!sheetName) return "No active crawl sheet found.";

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return "Crawl sheet \"" + sheetName + "\" not found.";

    // Save as shared filter
    props.setProperty("SNOWBALL_FILTER_GROUPS", JSON.stringify(groups));

    applyCrawlHighlight(sheet, groups);
    return "Highlight rule updated on \"" + sheetName + "\".";
  } catch (e) {
    return "Error: " + e.message;
  }
}

