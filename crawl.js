// ============================================================
// Crawl — shared infrastructure for the citation-crawl pipeline
//
// As of v22, this file holds ONLY the pieces the crawl pipeline (crawl_v2.js)
// genuinely depends on: sheet layout constants, sheet setup, row builders,
// S2 fetch/backoff helpers, the seed finder used by the panel's optional
// hand-picked-seed feature, and shared bookkeeping (matched-cites tracking,
// abstract notes). The original v1 forward/backward orchestration
// (crawlBatchTrigger, runCrawlLoop, runBackwardPass, startCrawl, resumeCrawl,
// the old hard-veto filter, cancelCrawl/trigger management, and the OpenAlex
// candidate path only that orchestration used) has been archived verbatim
// to archive/crawl_v1_full.js — preserved for reference/restoration, but no
// longer part of what's pushed to Apps Script (skipSubdirectories/.claspignore
// exclude archive/). Removing it here was confirmed only after tracing every
// v2/v3 call site (direct and transitive) back into this file, since roughly
// half of it turned out to be shared infrastructure rather than v1-specific.
// ============================================================

// Cols 1-14 are "owned" data/formula columns.
// Col 16 is "In-Sheet Links" — a MAP formula column, not in CRAWL_HEADERS.
// Term-helper columns start at CRAWL_FIRST_DETAIL_COL (17).
const CRAWL_HEADERS  = ["Depth","Crawled","Year","Title","Authors","Type","Venue","Abstract","ID","Cited By","Filter Match","Review","Found From","Matched Cites","Direction"];
const CRAWL_NUM_COLS = CRAWL_HEADERS.length; // 15

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
  FILTER_MATCH:  11, // pure boolean — TRUE iff in scope (every positive group matched, no exclude-mode NOT group tripped), independent of REVIEW_FLAG
  REVIEW_FLAG:   12, // pure boolean — TRUE iff FILTER_MATCH=TRUE AND a review-mode NOT group also tripped (needs human triage)
  FOUND_FROM:    13,  // parent paper ID (written at crawl time)
  MATCHED_CITES: 14, // in-sheet child IDs (written by updateCrawlMatchedCites)
  DIRECTION:     15  // "V" = venue sweep, "K" = keyword pass, "B" = backward pass, "F" = forward pass
};

// Col 16: In-Sheet Links — MAP formula; sits after Direction and before term helpers.
// Counts how many papers in the sheet list this paper as their Found From parent (col M).
const CRAWL_IN_SHEET_LINKS_COL = CRAWL_COL.DIRECTION + 1; // 16
const CRAWL_IN_SHEET_LINKS_FORMULA =
  '=MAP(A2:A,I2:I,LAMBDA(a,id,' +
  'IF(ROW(a)=2,"In-Sheet Links",' +
  'IF(a="","",COUNTIF(M$3:M,id)))))';

// Term-helper columns written by the highlight applier start here.
const CRAWL_FIRST_DETAIL_COL = CRAWL_IN_SHEET_LINKS_COL + 1; // 17

// Column letters for FILTER_MATCH (col 11 = K) and REVIEW_FLAG (col 12 = L)
// used in CF formulas. NOTE: these, and the hardcoded "M" in
// CRAWL_IN_SHEET_LINKS_FORMULA above (Found From), are literal column
// letters, not derived from CRAWL_COL — if either column's position ever
// moves again, these must be updated by hand alongside it.
const CRAWL_FILTER_MATCH_COL_LETTER = "K";
const CRAWL_REVIEW_FLAG_COL_LETTER  = "L";

// Default MAP formulas — evaluate to FALSE until a real filter is applied.
const CRAWL_DEFAULT_FILTER_FORMULA =
  '=MAP(A2:A,H2:H,LAMBDA(a,h,IF(ROW(a)=2,"Filter Match",IF(a<>"",FALSE,""))))';
const CRAWL_DEFAULT_REVIEW_FORMULA =
  '=MAP(A2:A,H2:H,LAMBDA(a,h,IF(ROW(a)=2,"Review",IF(a<>"",FALSE,""))))';

const CRAWL_ROW_HEIGHT = 60;

// Semantic Scholar back-off delays (ms) applied on HTTP 429 responses.
// Sequence: 3 s → 10 s → 30 s (then give up / surface error).
const S2_BACKOFF_MS = [3000, 10000, 30000];

// Row-1 status indicator — sits in the Year + Title columns of the header row,
// which are otherwise unused at row 1.
const CRAWL_STATUS_LABEL_COL = 3;  // "Crawl Status" label (Year col, row 1)
const CRAWL_STATUS_VALUE_COL = 4;  // status value with colour coding (Title col, row 1)

// How many consecutive fetch failures on the SAME paper (across separate
// crawl sessions — see markFetchFailure below) are tolerated before giving
// up on it, rather than leaving it queued forever.
const CRAWL_FETCH_FAILURE_MAX_RETRIES = 3;

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

// Fetches a Semantic Scholar URL, silently retrying with exponential
// back-off (3 s → 10 s → 30 s) on 429 (rate limit) AND on 5xx (transient
// server error) — confirmed live: complex bulk-search queries occasionally
// return a bare "Internal Server Error" with no rate-limit signal at all,
// and an identical retry moments later succeeds. Before this, only 429 was
// retried, so a 500 looked identical to "not 200" to every caller —
// several of which (s2BulkSearch in particular) treated any non-200
// response as "genuinely zero results, nothing more to fetch", silently
// truncating an exhaustive search partway through. Used for background
// crawl-loop and seed-resolution calls.
function s2Fetch(url) {
  var opts = getS2FetchOptions();
  var resp;
  for (var i = 0; i <= S2_BACKOFF_MS.length; i++) {
    resp = UrlFetchApp.fetch(url, opts);
    var code = resp.getResponseCode();
    var retryable = code === 429 || code >= 500;
    if (!retryable || i === S2_BACKOFF_MS.length) break;
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

  // Row 2 headers — three segments because cols 11, 12, and 16 are MAP formulas, not static text.
  // Cols 1-10: static text (Depth … Cited By)
  sheet.getRange(2, 1, 1, CRAWL_COL.CITED_BY)
    .setValues([CRAWL_HEADERS.slice(0, CRAWL_COL.CITED_BY)])
    .setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  // Col 11: Filter Match MAP formula
  sheet.getRange(2, CRAWL_COL.FILTER_MATCH)
    .setFormula(CRAWL_DEFAULT_FILTER_FORMULA)
    .setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  // Col 12: Review MAP formula
  sheet.getRange(2, CRAWL_COL.REVIEW_FLAG)
    .setFormula(CRAWL_DEFAULT_REVIEW_FORMULA)
    .setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  // Cols 13-15: static text (Found From, Matched Cites, Direction)
  sheet.getRange(2, CRAWL_COL.FOUND_FROM, 1, 3)
    .setValues([CRAWL_HEADERS.slice(CRAWL_COL.FOUND_FROM - 1)])
    .setFontWeight("bold").setBackground("#4285f4").setFontColor("white");
  // Col 16: In-Sheet Links MAP formula
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
  sheet.setColumnWidth(CRAWL_COL.FILTER_MATCH,      90);
  sheet.setColumnWidth(CRAWL_COL.REVIEW_FLAG,       90);
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

  // CF rule: highlight entire row green when Filter Match = TRUE.
  // (v2's applyCrawlV2Highlight replaces this with a green/orange pair the
  // first time "Apply Highlight Rule" runs — this is just the placeholder
  // that's correct before any real filter groups exist.)
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

// Returns the first row where Crawled = FALSE (skipping any row numbers in
// skipRows, if given), or null if none. skipRows lets a caller move past a
// row that just failed and is being left queued for a LATER session to
// retry, without the same while-loop immediately re-selecting it and
// burning the whole time budget retrying one broken row (see
// markFetchFailure / CRAWL_FETCH_FAILURE_MAX_RETRIES).
function findNextUncrawled(sheet, skipRows) {
  const lastRow = getCrawlLastDataRow(sheet);
  if (lastRow < 3) return null;
  const data = sheet.getRange(3, 1, lastRow - 2, CRAWL_NUM_COLS).getValues();
  for (var i = 0; i < data.length; i++) {
    var sheetRow = i + 3;
    if (skipRows && skipRows.has(sheetRow)) continue;
    if (data[i][CRAWL_COL.CRAWLED - 1] === false) {
      return {
        sheetRow: sheetRow,
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
    "",               // Review — spilled from row 2 formula
    parentId || "",   // Found From
    "",               // Matched Cites — populated by updateCrawlMatchedCites
    dir || "F"        // Direction
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

// ============================================================
// Term matching helper (shared with the consolidated filter in crawl_v2.js)
// ============================================================

// Escapes regex metacharacters so a filter term can be dropped into a
// RegExp/REGEXMATCH pattern literally.
function escapeRegExpTerm(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Builds the inner alternation for a term-match pattern, tolerant of a
// trailing regular plural ("s"/"es") and, for terms ending in "child", the
// irregular "...children" plural too — shared by termsAnyMatchV2,
// firstMatchingTerm, and buildTermFormulaV2 (crawl_v2.js) so plural
// handling lives in one place instead of three independently-patched
// \bTERM\b builders (the exact "more than one implementation" pattern
// this project's filter consolidation was trying to avoid elsewhere).
// Caller wraps the return value in \b...\b — only the TRAILING boundary
// needs plural tolerance; the leading \b is intentionally untouched
// (verified against Thai/Dubai/Mumbai for "AI" — no false-positive risk
// there, and this change doesn't touch that side at all).
//
// escapeFn: caller-supplied escaper, since JS-RegExp escaping and the
// Sheets-formula variant (which also needs `"` doubled) differ slightly.
function buildPluralAwareTermPattern(term, escapeFn) {
  var t = (term || '').toLowerCase();
  var mainAlt = escapeFn(t) + '(?:e?s)?';
  if (!/child$/i.test(t)) return mainAlt;
  var stem = t.slice(0, t.length - 'child'.length); // '' if term is exactly "child"
  return '(?:' + mainAlt + '|' + escapeFn(stem) + 'children)';
}

// ============================================================
// Fetch candidates
// ============================================================

// Forward: papers that cite the given paper (Semantic Scholar)
function fetchForwardCandidates(id, title) {
  return s2GetCitations(id, title); // from snowball.js
}

// Turns a fetchOpenAlexAbstractV2() result (crawl_v2.js) into a
// human-readable note for the Abstract cell — only called when Semantic
// Scholar itself had no abstract.
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

// Prefixes a paper's Title cell with a visible failure marker when its
// candidate fetch (citations or references) throws, tracking a retry count
// so a persistent failure eventually gives up rather than looping forever,
// while a transient one (e.g. a 429 that's since cleared) gets a bounded
// number of automatic retries across separate crawl sessions instead of
// becoming a permanent corrupted title (v22 §7 bug: a rate-limit message
// was being written into Title and left there forever, with the row marked
// done and never retried).
//
// Returns true if the caller should leave the row queued for a later retry
// (Crawled=FALSE), false if it should give up (Crawled=TRUE). Deliberately
// reuses the Title column rather than adding a new one, since a new column
// would shift every fixed column position (CRAWL_COL.*) and the
// term-helper/filter-formula columns after it, corrupting any crawl sheet
// created before this change if it's ever resumed.
function markFetchFailure(sheet, sheetRow, e) {
  var cell     = sheet.getRange(sheetRow, CRAWL_COL.TITLE);
  var curTitle = String(cell.getValue());
  var match    = curTitle.match(/^⚠ \[fetch failed (\d+)x:[^\]]*\]\s*/);
  var attempt  = match ? parseInt(match[1]) + 1 : 1;
  var baseTitle = curTitle.replace(/^⚠ \[fetch failed \d+x:[^\]]*\]\s*/, '');
  var shouldRetry = attempt < CRAWL_FETCH_FAILURE_MAX_RETRIES;
  var suffix = shouldRetry ? 'will retry' : 'giving up';
  cell.setValue('⚠ [fetch failed ' + attempt + 'x: ' + e.message.slice(0, 60) + ' — ' + suffix + '] ' + baseTitle);
  return shouldRetry;
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

  // --- Format detection ---
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
