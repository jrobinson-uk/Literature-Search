// ============================================================
// Crawl v2 — parallel pipeline: Keyword pass → Backward pass → Forward pass
//
// Built alongside (not replacing) crawl.js's forward/backward crawler, per
// the "Scholar Search Tool v2 pipeline" brief validated with the user on
// 2026-08-14. Reuses the v1 sheet layout (CRAWL_HEADERS/CRAWL_COL), row
// builder (crawlRowFromS2), S2/OpenAlex fetch helpers, term-formula
// machinery, and log integration from crawl.js/snowball.js/log.js
// unmodified — only the phase orchestration, keyword-pass generation,
// dedup, and NOT-group handling are new. v1's own crawlBatchTrigger /
// CRAWL_* properties are completely untouched.
//
// Entirely separate script-property namespace (CRAWL2_*) and trigger
// (crawlV2BatchTrigger), so a v1 crawl and a v2 crawl can run at the same
// time in the same spreadsheet without colliding — including a dedicated
// CRAWL2_FILTER_GROUPS key rather than reusing the shared
// SNOWBALL_FILTER_GROUPS property that v1 Crawl and Snowball already share,
// so an in-progress v2 run's filter can't be clobbered by unrelated panel
// activity elsewhere.
//
// Validated decisions this build reflects (see chat for full validation
// write-up against the brief):
//   - Phases run strictly sequentially: keyword → backward → forward.
//   - Keyword pass REPLACES the 10 hand-picked seeds as the default seed
//     source; hand-picked seeds remain addable as an explicit supplement.
//   - NOT-group hits are score-demoted (included + flagged), not a hard
//     veto — a paper passing every positive group is kept even if a
//     NOT-group also matched.
//   - Abstract resolution stays Semantic Scholar + OpenAlex only (no third
//     source), but persists longer: more backoff attempts per lookup, plus
//     a final retry sweep over unresolved rows before the run completes.
//   - Near-miss (2-of-3 group) tier: skipped for this build.
//   - Keyword pass uses Semantic Scholar's free /paper/search endpoint
//     (already integrated elsewhere in this project), targeting ~200 seeds
//     by default.
// ============================================================

// Literal values, not references to crawl.js's CRAWL_TIME_LIMIT_MS /
// CRAWL_MAX_CONSEC_FAILURES — Apps Script concatenates every .js file into
// one execution context, and top-level `const` initialization across files
// depends on file load order, which isn't something to rely on. Same
// values as v1, kept in sync manually; duplicating two numbers here is a
// far smaller risk than a load-order-dependent ReferenceError that would
// silently break onOpen() (and with it, the whole custom menu) if the
// files ever load in the "wrong" order.
const CRAWL2_TIME_LIMIT_MS       = 4 * 60 * 1000;
const CRAWL2_MAX_CONSEC_FAILURES = 3;

// Longer retry effort than v1's OPENALEX_BACKOFF_MS — validated response to
// "persist with the search for longer before moving on" rather than adding
// a third abstract source.
const OPENALEX_BACKOFF_MS_V2 = [1000, 2000, 4000, 8000, 15000, 30000];

// v3 additions — off by default. §7 sign-off (2026-08-20): extend v2 in
// place with these as opt-in config, not a separate v3 codebase; a run
// with all of them left off reproduces today's (v19) behaviour exactly.
const PHASE0_YEAR_FROM_DEFAULT = 2023;
const PHASE0_YEAR_TO_DEFAULT   = 2026;
// A-priori venue list from the v3 brief — added regardless of whether they
// were represented in the v19 run. The other ~48 venues the brief derived
// from v19 (those with ≥5 matches) aren't available here since that export
// wasn't provided; the panel's venue list is a free-text field the user can
// extend with them directly.
const PHASE0_VENUES_DEFAULT = [
  'IDC', 'WiPSCE', 'Koli Calling', 'TOCE', 'ICER V.2', 'SIGCSE V.2'
];
// S2's bulk-search endpoint returns up to this many rows per page,
// paginated via a continuation token (confirmed against the live API,
// not assumed) — used by both Phase 0 (venue enumeration) and Phase 1's
// paginated alt-config.
const S2_BULK_PAGE_SIZE = 1000;
// How many venue names to pack into one bulk-search call's comma-separated
// `venue` filter — S2 matches each fuzzily against canonical venue names
// (confirmed: "SIGCSE" alone matched "Technical Symposium on Computer
// Science Education"), so batching keeps the URL short without needing
// exact venue-string matches.
const PHASE0_VENUE_BATCH_SIZE = 10;
// ============================================================
// Timestamp / duration / progress header cells (v2-only — row 1, columns
// 5-10, between the existing Crawl Status pair at cols 3-4 and the filter-
// group headers that start at CRAWL_FIRST_DETAIL). Not part of CRAWL_COL
// (the per-row data columns), so this doesn't touch sheet resumability.
// ============================================================
const CRAWL2_STARTED_LABEL_COL  = 5;
const CRAWL2_STARTED_VALUE_COL  = 6;
const CRAWL2_DURATION_LABEL_COL = 7;
const CRAWL2_DURATION_VALUE_COL = 8;
const CRAWL2_PROGRESS_LABEL_COL = 9;
const CRAWL2_PROGRESS_VALUE_COL = 10;

// Fixed pipeline order, purely for rendering "Phase X/6: <name>" in the
// header — not read anywhere for control flow (crawlV2BatchTrigger's own
// if/else chain remains the source of truth for what actually runs next).
const CRAWL2_PHASE_ORDER = [
  { key: 'venue',     label: 'Venue sweep' },
  { key: 'keyword',   label: 'Keyword pass' },
  { key: 'backward',  label: 'Backward pass' },
  { key: 'forward',   label: 'Forward pass' },
  { key: 'backward2', label: 'Second backward pass' },
  { key: 'sweep',     label: 'Final abstract sweep' }
];

function formatDurationV2(ms) {
  if (ms < 0) ms = 0;
  var totalSeconds = Math.floor(ms / 1000);
  var days  = Math.floor(totalSeconds / 86400);
  var hours = Math.floor((totalSeconds % 86400) / 3600);
  var mins  = Math.floor((totalSeconds % 3600) / 60);
  var secs  = totalSeconds % 60;
  if (days  > 0) return days + 'd ' + hours + 'h';
  if (hours > 0) return hours + 'h ' + mins + 'm';
  if (mins  > 0) return mins + 'm ' + secs + 's';
  return secs + 's';
}

function formatTimestampV2(date) {
  var pad = function(n) { return n.toString().padStart(2, '0'); };
  return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate()) +
    ' ' + pad(date.getHours()) + ':' + pad(date.getMinutes());
}

// Writes the three label cells once (crawl creation only) — values are
// filled in by updateCrawlV2Timing as the crawl actually runs.
function setupCrawlV2TimingHeaders(sheet) {
  sheet.getRange(1, CRAWL2_STARTED_LABEL_COL)
    .setValue('Started').setFontWeight('bold').setFontColor('#555555').setFontSize(10);
  sheet.getRange(1, CRAWL2_DURATION_LABEL_COL)
    .setValue('Duration').setFontWeight('bold').setFontColor('#555555').setFontSize(10);
  sheet.getRange(1, CRAWL2_PROGRESS_LABEL_COL)
    .setValue('Progress').setFontWeight('bold').setFontColor('#555555').setFontSize(10);
  sheet.setColumnWidth(CRAWL2_STARTED_VALUE_COL,  110);
  sheet.setColumnWidth(CRAWL2_DURATION_VALUE_COL,  70);
  sheet.setColumnWidth(CRAWL2_PROGRESS_VALUE_COL, 220);
}

// Refreshes Duration + Progress on every batch-trigger firing (Started is
// written once, at crawl creation, and never changes). Wall-clock elapsed
// since crawl start, including any time spent waiting for a manual Resume —
// "how long has this crawl been going" rather than pure execution time.
function updateCrawlV2Timing(sheet, phase) {
  var props    = PropertiesService.getScriptProperties();
  var startIso = props.getProperty('CRAWL2_START_TIME');
  if (startIso) {
    var elapsed = Date.now() - new Date(startIso).getTime();
    sheet.getRange(1, CRAWL2_DURATION_VALUE_COL).setValue(formatDurationV2(elapsed));
  }
  var idx = -1;
  for (var i = 0; i < CRAWL2_PHASE_ORDER.length; i++) {
    if (CRAWL2_PHASE_ORDER[i].key === phase) { idx = i; break; }
  }
  var progressText = (idx >= 0)
    ? 'Phase ' + (idx + 1) + '/' + CRAWL2_PHASE_ORDER.length + ': ' + CRAWL2_PHASE_ORDER[idx].label
    : (phase === 'complete' ? 'Complete' : (phase || ''));
  sheet.getRange(1, CRAWL2_PROGRESS_VALUE_COL).setValue(progressText);
}

// ============================================================
// Trigger / status management (own namespace, mirrors crawl.js)
// ============================================================

function deleteCrawlV2Trigger() {
  var props = PropertiesService.getScriptProperties();
  var id    = props.getProperty('CRAWL2_TRIGGER_ID');
  if (!id) return;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getUniqueId() === id) ScriptApp.deleteTrigger(t);
  });
  props.deleteProperty('CRAWL2_TRIGGER_ID');
}

function isCrawlV2Running() {
  var id = PropertiesService.getScriptProperties().getProperty('CRAWL2_TRIGGER_ID');
  if (!id) return false;
  return ScriptApp.getProjectTriggers().some(function(t) { return t.getUniqueId() === id; });
}

function createCrawlV2Trigger() {
  deleteCrawlV2Trigger();
  var trigger = ScriptApp.newTrigger('crawlV2BatchTrigger')
    .timeBased().everyMinutes(1).create();
  PropertiesService.getScriptProperties()
    .setProperty('CRAWL2_TRIGGER_ID', trigger.getUniqueId());
}

// Menu handler: stops the running v2 crawl's trigger without wiping its
// state, so Resume v2 Crawl still works afterward.
function cancelCrawlV2() {
  var ui = SpreadsheetApp.getUi();
  if (!isCrawlV2Running()) {
    ui.alert('No v2 crawl is currently running.');
    return;
  }
  var props     = PropertiesService.getScriptProperties();
  var sheetName = props.getProperty('CRAWL2_ACTIVE_SHEET');
  var response  = ui.alert('Cancel Crawl v2',
    'Stop the currently running v2 crawl ("' + sheetName + '")?\n\n' +
    'Progress so far is kept — you can still click "Resume v2 Crawl" later.',
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  deleteCrawlV2Trigger();
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var logRow = parseInt(props.getProperty('CRAWL2_LOG_ROW') || '0') || 0;
  if (sheet)  setCrawlStatus(sheet, 'Cancelled');   // reused from crawl.js — identical status cell layout
  if (logRow) updateLogRow(logRow, 'Cancelled');
  ui.alert('v2 crawl cancelled.');
}

// Read-only progress check for whichever phase is currently active — works
// on ANY running crawl, including ones started before the Started/
// Duration/Progress header cells existed, by re-deriving the same
// eligibility/queue logic each phase uses internally (mirrors
// runBackwardPassV2's own paperIds-building loop for backward/backward2).
// Never writes to properties or the sheet — safe to run mid-crawl.
function debugCrawlV2Progress() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var sheetName = props.getProperty('CRAWL2_ACTIVE_SHEET');
  if (!sheetName) { ui.alert('No active v2 crawl found.'); return; }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) { ui.alert('Crawl sheet "' + sheetName + '" not found.'); return; }

  var phase       = props.getProperty('CRAWL2_PHASE') || 'venue';
  var batch       = parseInt(props.getProperty('CRAWL2_BATCH_NUM') || '1');
  var maxDepth    = parseInt(props.getProperty('CRAWL2_MAX_DEPTH') || '2');
  var groups      = JSON.parse(props.getProperty('CRAWL2_FILTER_GROUPS') || '[]');
  var guardPhrases = JSON.parse(props.getProperty('CRAWL2_GUARD_PHRASES') || '[]');
  var minutesPerBatch = CRAWL2_TIME_LIMIT_MS / 60000;

  var phaseIdx = -1;
  for (var pi = 0; pi < CRAWL2_PHASE_ORDER.length; pi++) {
    if (CRAWL2_PHASE_ORDER[pi].key === phase) { phaseIdx = pi; break; }
  }
  var lines = [
    'Phase: ' + (phaseIdx >= 0
      ? ('Phase ' + (phaseIdx + 1) + '/' + CRAWL2_PHASE_ORDER.length + ' — ' + CRAWL2_PHASE_ORDER[phaseIdx].label)
      : phase),
    'Batch: ' + batch + ' (each batch runs up to ~' + minutesPerBatch + ' min before re-triggering)'
  ];

  if (phase === 'backward' || phase === 'backward2') {
    var propPrefix = (phase === 'backward2') ? 'CRAWL2_BACKWARD2' : 'CRAWL2_BACKWARD';
    var idx      = parseInt(props.getProperty(propPrefix + '_IDX')      || '0');
    var examined = parseInt(props.getProperty(propPrefix + '_EXAMINED') || '0');
    var kept     = parseInt(props.getProperty(propPrefix + '_KEPT')     || '0');

    var lastRow = getCrawlLastDataRow(sheet);
    var total = 0;
    if (lastRow >= CRAWL_FIRST_DATA_ROW) {
      var data = sheet.getRange(CRAWL_FIRST_DATA_ROW, 1, lastRow - CRAWL_FIRST_DATA_ROW + 1, CRAWL_NUM_COLS).getValues();
      data.forEach(function(row) {
        var id = String(row[CRAWL_COL.ID - 1] || '').trim();
        if (!id) return;
        var depth = parseInt(row[CRAWL_COL.DEPTH - 1]) || 0;
        if (depth >= maxDepth) return;
        var title    = String(row[CRAWL_COL.TITLE - 1]    || '');
        var abstract = String(row[CRAWL_COL.ABSTRACT - 1] || '');
        if (!jsMatchesFilterV2(title + ' ' + abstract, groups, guardPhrases).expand) return;
        total++;
      });
    }
    lines.push('Source papers examined: ' + idx + ' / ' + total);
    lines.push('References examined so far: ' + examined + ' (kept: ' + kept + ')');
    if (idx > 0 && total > idx) {
      var minutesSoFar   = batch * minutesPerBatch;
      var minutesPerSource = minutesSoFar / idx;
      var remainingMin   = Math.round(minutesPerSource * (total - idx));
      lines.push('Rough estimate: ~' + remainingMin + ' more minute(s) at the current pace — ' +
        'very approximate, since it depends on how many references each remaining paper has.');
    }
  } else if (phase === 'forward') {
    lines.push('Papers still queued (Crawled=FALSE): ' + countUncrawled(sheet));
  } else if (phase === 'keyword') {
    var subQueries = buildKeywordSubQueries(groups);
    var subQueryIdx = parseInt(props.getProperty('CRAWL2_KEYWORD_SUBQUERY_IDX') || '0');
    var currentTerm = subQueries[subQueryIdx] ? subQueries[subQueryIdx].term : '(none)';
    lines.push('Sub-query: ' + (subQueryIdx + 1) + ' / ' + subQueries.length + ' — term: "' + currentTerm + '"');
    lines.push('Matches collected: ' + (props.getProperty('CRAWL2_KEYWORD_COLLECTED') || '0'));
    lines.push('Candidates examined: ' + (props.getProperty('CRAWL2_KEYWORD_RESULTS_SEEN') || '0'));
    lines.push('Pages fetched: ' + (props.getProperty('CRAWL2_KEYWORD_PAGES_FETCHED') || '0'));
    var subQueryErrs = parseInt(props.getProperty('CRAWL2_KEYWORD_SUBQUERY_ERRORS') || '0');
    var totalErrs    = parseInt(props.getProperty('CRAWL2_KEYWORD_TOTAL_ERRORS')    || '0');
    if (subQueryErrs > 0) {
      lines.push('⚠ Current sub-query has hit ' + subQueryErrs + ' transient S2 error(s) so far — retrying.');
    }
    lines.push('Total transient S2 errors this pass: ' + totalErrs);
    var retriedTermsSoFar = JSON.parse(props.getProperty('CRAWL2_KEYWORD_RETRIED_TERMS') || '[]');
    if (retriedTermsSoFar.length > 0) {
      lines.push('Sub-queries that needed a retry so far: ' + retriedTermsSoFar.join(', '));
    }
  } else if (phase === 'venue') {
    lines.push('Seed papers collected so far: ' + (props.getProperty('CRAWL2_VENUE_COUNT') || '0'));
  } else if (phase === 'sweep') {
    lines.push('Abstracts recovered so far: ' + (props.getProperty('CRAWL2_SWEEP_RECOVERED') || '0'));
  }

  ui.alert('Citation Crawl Progress', lines.join('\n'), ui.ButtonSet.OK);
}

// ============================================================
// Sheet setup — reuses v1's layout entirely
// ============================================================

function setupCrawlV2Sheet(sheet, seedLabel) {
  setupCrawlSheet(sheet, 'forward', seedLabel); // identical column layout, placeholder formula, and CF rules to v1
  sheet.getRange(1, 1).setValue('Pipeline crawl (v2: keyword → backward → forward)');
  // Filter Match stays boolean TRUE/FALSE (same as v1) — the placeholder is
  // FALSE-for-all until a real filter is applied, same as v1, so nothing
  // else needs overriding here. applyCrawlV2Highlight below replaces v1's
  // single green row rule with a refined green/orange pair once real
  // filter groups exist.
  setupCrawlV2TimingHeaders(sheet);
}

// ============================================================
// Dedup — ID *and* normalised-title, catching preprint/published-version
// duplicates that share no ID but are the same paper (the brief's dedup
// gap: currently keyed by graph-API ID only).
// ============================================================

function normalizeTitleV2(title) {
  return (title || '')
    .toLowerCase()
    .replace(/^⚠ \[[^\]]*\]\s*/, '')  // strip any existing marker prefix first
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

// getCrawlV2ExistingKeys reconstructs ID + normalised-title keys from the
// sheet itself, but DOI isn't stored in any column (the sheet's ID
// convention is W<mag>/S2:<hash>, never a DOI), so a DOI-based duplicate
// can't be reconstructed retroactively by re-scanning the sheet the way ID/
// title can. Instead, DOIs seen so far THIS crawl are accumulated in a
// script property as candidates are written, across all phases and
// sessions of the same crawl — catching cross-phase duplicates like
// "App Planner via both the venue pass and the keyword pass" (v22 §8) even
// when the two occurrences' titles differ enough that normalised-title
// dedup wouldn't catch them, as long as both carry the same DOI.
function loadSeenDois() {
  try { return new Set(JSON.parse(PropertiesService.getScriptProperties().getProperty('CRAWL2_SEEN_DOIS') || '[]')); }
  catch (e) { return new Set(); }
}
function saveSeenDois(doiSet) {
  // Cap stored size defensively — PropertiesService has a per-property
  // limit; a crawl accumulating tens of thousands of distinct DOIs is far
  // beyond what this project's runs have ever produced, but fail safe
  // (stop persisting further growth) rather than throw mid-crawl.
  var arr = Array.from(doiSet);
  if (arr.length > 20000) return;
  PropertiesService.getScriptProperties().setProperty('CRAWL2_SEEN_DOIS', JSON.stringify(arr));
}

function getCrawlV2ExistingKeys(sheet) {
  var lastRow = getCrawlLastDataRow(sheet);
  var ids     = new Set();
  var titles  = new Set();
  var dois    = loadSeenDois();
  if (lastRow < CRAWL_FIRST_DATA_ROW) return { ids: ids, titles: titles, dois: dois };
  var numRows   = lastRow - CRAWL_FIRST_DATA_ROW + 1;
  var idVals    = sheet.getRange(CRAWL_FIRST_DATA_ROW, CRAWL_COL.ID,    numRows, 1).getValues().flat();
  var titleVals = sheet.getRange(CRAWL_FIRST_DATA_ROW, CRAWL_COL.TITLE, numRows, 1).getValues().flat();
  idVals.forEach(function(v) { if (v) ids.add(v); });
  titleVals.forEach(function(v) {
    var norm = normalizeTitleV2(v);
    if (norm) titles.add(norm);
  });
  return { ids: ids, titles: titles, dois: dois };
}

// True if this candidate is a duplicate of something already in `existing`
// (ID, normalised title, or DOI) — checked once per candidate across every
// phase, so the same three-way dedup applies everywhere consistently.
function isDuplicateCandidateV2(existing, id, normTitle, doi) {
  if (existing.ids.has(id)) return true;
  if (normTitle && existing.titles.has(normTitle)) return true;
  if (doi && existing.dois.has(doi)) return true;
  return false;
}

// Records a kept candidate's keys into `existing` (in-memory, for the rest
// of this batch) AND persists its DOI (if any) into the cross-session
// CRAWL2_SEEN_DOIS store, since that's the one key that can't be
// reconstructed later just by re-scanning the sheet.
function rememberCandidateV2(existing, id, normTitle, doi, persistDoi) {
  existing.ids.add(id);
  if (normTitle) existing.titles.add(normTitle);
  if (doi) {
    existing.dois.add(doi);
    // Only persist to the cross-session store for genuine matches. With a
    // phase's matchesOnly=false, FALSE candidates get written to the sheet
    // too (for audit) and still flow through here for in-memory dedup —
    // but persisting THEIR DOIs as well is what blew PropertiesService's
    // 500KB total quota on a real run (15k+ recorded non-matches, each
    // triggering a full saveSeenDois rewrite of the growing array). FALSE
    // candidates don't need remembering across phases/sessions: the
    // sheet's own Title/ID columns already prevent re-writing them once
    // getCrawlV2ExistingKeys re-scans it on the next phase call.
    if (persistDoi !== false) saveSeenDois(existing.dois);
  }
}

// ============================================================
// Consolidated tri-state filter (v22 §2/§3)
//
// One implementation, called identically from every phase (venue, keyword,
// backward, forward) AND mirrored by the sheet's own live Filter Match
// formula below (buildFilterMatchFormulaV2 / buildTermFormulaV2) — the
// brief's diagnosed root cause for "the filter behaves differently per
// phase" was more than one implementation computing (or approximating) the
// same thing; this keeps it to one JS function plus one formula-builder
// that both encode the exact same rules.
//
//   TRUE   — every positive group matched, no NOT group tripped
//   FALSE  — at least one positive group failed to match (checked first,
//            regardless of any NOT group — never overridden by a NOT hit)
//   REVIEW — every positive group matched, AND a NOT group whose notMode
//            is "review" tripped. If BOTH an "exclude"-mode group and a
//            "review"-mode group trip, exclude wins (FALSE) — per §2's
//            explicit rule.
//
// REVIEW rows are harvested (written, flagged) but never queued for
// backward/forward expansion — callers gate on state, not a boolean.
//
// guardPhrases (global, not per-group): phrases that suppress an otherwise-
// matching term when its only occurrence is inside one of them (e.g.
// "Scratch" inside "from scratch"). Implemented by masking every
// guardPhrase occurrence out of the text before term-matching runs, both
// here and in the sheet-formula term-helper columns, so a term whose ONLY
// occurrence was inside a guarded phrase doesn't match either way, while a
// separate, real occurrence elsewhere in the same text still does.
// ============================================================

function maskGuardPhrases(text, guardPhrases) {
  if (!guardPhrases || guardPhrases.length === 0) return text;
  guardPhrases.forEach(function(phrase) {
    var p = (phrase || '').trim();
    if (!p) return;
    var re = new RegExp('\\b' + escapeRegExpTerm(p.toLowerCase()) + '\\b', 'gi'); // escapeRegExpTerm from crawl.js
    text = text.replace(re, function(m) { return new Array(m.length + 1).join(' '); });
  });
  return text;
}

function parseGroupTermList(group) {
  if (!group.terms || !group.terms.trim()) return [];
  return group.terms.split(",")
    .map(function(t) { return t.trim().replace(/^["']|["']$/g, "").trim().toLowerCase(); })
    .filter(function(t) { return t.length > 0; });
}

function termsAnyMatchV2(text, group) {
  var terms = parseGroupTermList(group);
  if (terms.length === 0) return false;
  return terms.some(function(t) {
    return new RegExp('\\b' + buildPluralAwareTermPattern(t, escapeRegExpTerm) + '\\b', 'i').test(text); // buildPluralAwareTermPattern from crawl.js
  });
}

// First matching term in a group, for the Flag Reason note on a REVIEW row
// — so a reviewer doesn't have to re-derive which term triggered it.
function firstMatchingTerm(text, group) {
  var terms = parseGroupTermList(group);
  for (var i = 0; i < terms.length; i++) {
    if (new RegExp('\\b' + buildPluralAwareTermPattern(terms[i], escapeRegExpTerm) + '\\b', 'i').test(text)) return terms[i];
  }
  return null;
}

// Returns { state, isMatch, expand, flagGroupIndex, flagTerm }.
//   isMatch — true for TRUE or REVIEW (worth keeping/harvesting)
//   expand  — true only for TRUE (eligible for backward/forward expansion)
//   flagGroupIndex/flagTerm — set only for REVIEW, which NOT group + term
//     triggered it (for the Filter Match cell note)
function jsMatchesFilterV2(rawText, groups, guardPhrases) {
  var text = maskGuardPhrases((rawText || "").toLowerCase(), guardPhrases);
  var positive = groups.filter(function(g) { return !g.not; });
  var negative = groups.filter(function(g) { return g.not; });

  var positiveOk = positive.every(function(g) {
    if (!g.terms || !g.terms.trim()) return true;
    return termsAnyMatchV2(text, g);
  });
  if (!positiveOk) return { state: 'FALSE', isMatch: false, expand: false };

  // Exclude wins over review if both trip (§2's explicit rule).
  for (var i = 0; i < negative.length; i++) {
    var g = negative[i];
    if ((g.notMode || 'exclude') === 'exclude' && termsAnyMatchV2(text, g)) {
      return { state: 'FALSE', isMatch: false, expand: false };
    }
  }
  for (var j = 0; j < negative.length; j++) {
    var rg = negative[j];
    if ((rg.notMode || 'exclude') === 'review' && termsAnyMatchV2(text, rg)) {
      return {
        state: 'REVIEW', isMatch: true, expand: false,
        flagGroupIndex: groups.indexOf(rg), flagTerm: firstMatchingTerm(text, rg)
      };
    }
  }

  return { state: 'TRUE', isMatch: true, expand: true };
}

// ============================================================
// Filter Match + Review columns (two pure booleans) + row highlight
//
// Column K (Filter Match) and column L (Review) each hold a real Sheets
// boolean, computed by sheet formulas that mirror jsMatchesFilterV2's rules
// exactly: K = TRUE iff every positive group matched and no exclude-mode
// NOT group tripped (jsMatchesFilterV2's `isMatch` — in/out of scope,
// independent of any review flag); L = TRUE iff K is also TRUE AND a
// review-mode NOT group tripped (jsMatchesFilterV2's `state === 'REVIEW'`).
// Was a single tri-state text column (TRUE/FALSE/REVIEW) through v22; split
// into these two orthogonal booleans so an external check can read K alone
// for a definitive in-scope yes/no, with L as a separate human-triage flag.
// The per-term helper columns feeding both apply the same guardPhrases
// masking as the JS side (buildTermFormulaV2), so the live display can't
// drift from the write-time decision the way two separately-evolving
// implementations could.
// ============================================================

// Sheet-formula equivalent of maskGuardPhrases — nests REGEXREPLACE calls,
// one per guard phrase, each blanking out \b-bounded case-insensitive
// occurrences before the term match runs. `textExpr` is the raw
// LOWER(title&" "&abstract) expression to wrap.
function buildMaskedTextExprV2(textExpr, guardPhrases) {
  var expr = textExpr;
  (guardPhrases || []).forEach(function(phrase) {
    var p = (phrase || '').trim();
    if (!p) return;
    var esc = p.toLowerCase().replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/"/g, '""');
    expr = 'REGEXREPLACE(' + expr + ',"\\b' + esc + '\\b"," ")';
  });
  return expr;
}

// Composes escapeRegExpTerm (crawl.js) with the extra quote-doubling a
// Sheets formula string literal needs — the Sheets-specific half of
// buildPluralAwareTermPattern's escapeFn contract.
function escapeTermForSheetFormula(s) {
  return escapeRegExpTerm(s).replace(/"/g, '""'); // escapeRegExpTerm from crawl.js
}

// v2 variant of snowball.js's shared buildTermFormula, adding guardPhrases
// masking — not edited in place in snowball.js since that function is
// shared with the unrelated Snowball feature and v1.
//
// Anchored at CRAWL_FIRST_DATA_ROW rather than the header row — the header
// row now holds a plain static term label (written directly by the caller)
// and the row between that and this anchor holds a COUNTIF totals formula
// (also written by the caller); this array formula owns only the real data
// rows from CRAWL_FIRST_DATA_ROW down, so it can no longer double as its
// own header the way it did pre-totals-row (that trick would collide with
// the totals row now sitting in between).
function buildTermFormulaV2(term, titleColNum, abstractColNum, guardPhrases) {
  const titleLetter  = colToLetter(titleColNum); // from snowball.js
  const absLetter    = colToLetter(abstractColNum);
  const innerPattern = buildPluralAwareTermPattern(term, escapeTermForSheetFormula); // from crawl.js
  const maskedExpr   = buildMaskedTextExprV2('LOWER(t&" "&k)', guardPhrases);
  const anchor = CRAWL_FIRST_DATA_ROW;
  return '=MAP(A' + anchor + ':A,' + titleLetter + anchor + ':' + titleLetter + ',' +
         absLetter + anchor + ':' + absLetter + ',LAMBDA(a,t,k,' +
         'IF(a="","",' +
         'REGEXMATCH(' + maskedExpr + ',"\\b' + innerPattern + '\\b")' +
         ')))';
}

// Sheet formula for the totals row (CRAWL_TOTALS_ROW): counts how many rows
// in this column's own data range (CRAWL_FIRST_DATA_ROW downward) are TRUE
// — a plain COUNTIF, deliberately NOT part of the array formula above (a
// self-referential COUNTIF over a range the same array formula also
// produces would be circular), so it lives in a genuinely separate cell one
// row above that array's anchor.
function buildColumnTotalFormula(colNum) {
  const l = colToLetter(colNum); // from snowball.js
  return '=COUNTIF(' + l + CRAWL_FIRST_DATA_ROW + ':' + l + ',TRUE)';
}

// Shared building blocks for both the Filter Match and Review-flag formulas.
//
// Deliberately NOT a MAP/LAMBDA over one param per term any more (that's
// what the pre-totals-row version did, successfully, for years — but a live
// test straight after the totals-row deploy came back with Filter
// Match/Review blank for every single row, on a crawl with 5 filter groups
// / 37 total terms, i.e. a ~38-argument MAP/LAMBDA call). The per-term
// columns' OWN formulas (buildTermFormulaV2, only 3 params: a/t/k) kept
// working fine on the same crawl — the one thing that changed for THIS
// formula specifically, versus the term formulas, is the argument count.
// Never conclusively root-caused (no live Sheets access to inspect the
// actual cell state), but that's the one variable that survived every
// other elimination pass (anchor row, the removed ROW(a)=2 header-embed
// trick — both also changed in buildTermFormulaV2, which still works).
// Rather than ship a fix riding on an unconfirmed theory, this rebuilds the
// combination step on a completely different, much more boring mechanism:
// plain range arithmetic under ARRAYFORMULA (TRUE/FALSE coerce to 1/0, so
// "sum of a group's term columns > 0" is OR, and multiplying groups'
// results together is AND) — no LAMBDA, no per-term named parameters, so
// whatever the actual limit/quirk was, this can't hit it.
function buildFilterExprPartsV2(parsedGroups, firstDetailColNum) {
  const totalTerms = parsedGroups.reduce(function(s, g) { return s + g.terms.length; }, 0);
  if (totalTerms === 0) return null;

  function rangesForGroup(gi) {
    const startOffset = parsedGroups.slice(0, gi).reduce(function(s, pg) { return s + pg.terms.length; }, 0);
    return parsedGroups[gi].terms.map(function(t, ti) {
      const colNum = firstDetailColNum + startOffset + ti;
      const l = colToLetter(colNum); // from snowball.js
      return l + CRAWL_FIRST_DATA_ROW + ':' + l;
    });
  }

  // A group's own OR-across-terms, as a raw sum (>0 test applied by callers)
  // — parenthesized whenever it's more than one range so it composes safely
  // inside a larger arithmetic expression.
  function groupSumExpr(gi) {
    const ranges = rangesForGroup(gi);
    return ranges.length === 1 ? ranges[0] : ('(' + ranges.join('+') + ')');
  }

  const positiveIdx = parsedGroups.map(function(g, gi) { return g.not ? null : gi; })
    .filter(function(x) { return x !== null; });
  // Product of "(group sum > 0)" across every positive group — 1 only if
  // EVERY positive group had at least one term match (AND), 0 if any didn't.
  const positiveMatchExpr = positiveIdx.length === 0
    ? '1'
    : positiveIdx.map(function(gi) { return '(' + groupSumExpr(gi) + '>0)'; }).join('*');

  // Exclude/review only ever need "did ANY term across ALL such groups
  // match" (an OR of ORs is just one flat OR), so their terms are combined
  // into a single flat sum rather than kept per-group.
  function flatRanges(idxList) {
    var out = [];
    idxList.forEach(function(gi) { out = out.concat(rangesForGroup(gi)); });
    return out;
  }

  const excludeIdx = parsedGroups.map(function(g, gi) { return (g.not && (g.notMode || 'exclude') === 'exclude') ? gi : null; })
    .filter(function(x) { return x !== null; });
  const excludeRanges = flatRanges(excludeIdx);
  const excludeMatchExpr = excludeRanges.length === 0 ? 'FALSE' : ('(' + excludeRanges.join('+') + ')>0');

  const reviewIdx = parsedGroups.map(function(g, gi) { return (g.not && g.notMode === 'review') ? gi : null; })
    .filter(function(x) { return x !== null; });
  const reviewRanges = flatRanges(reviewIdx);
  const reviewMatchExpr = reviewRanges.length === 0 ? 'FALSE' : ('(' + reviewRanges.join('+') + ')>0');

  return { positiveMatchExpr: positiveMatchExpr, excludeMatchExpr: excludeMatchExpr, reviewMatchExpr: reviewMatchExpr };
}

// Pure boolean (split from the old tri-state text column into two orthogonal
// booleans — Filter Match + Review, see buildReviewFlagFormulaV2 below):
// TRUE iff every positive group matched AND no exclude-mode NOT group
// tripped. A review-mode NOT group hit does NOT affect this column any
// more — that's entirely the Review column's job now. This is exactly
// jsMatchesFilterV2's `isMatch` — "in scope", independent of whether it
// also needs human triage — so an external check can read this column
// alone for a definitive in/out-of-scope yes/no.
function buildFilterMatchFormulaV2(parsedGroups, firstDetailColNum) {
  const parts = buildFilterExprPartsV2(parsedGroups, firstDetailColNum);
  if (!parts) return null;
  // Header text and the totals-row COUNTIF are written separately by the
  // caller — this array owns only the real data rows, same as buildTermFormulaV2.
  return '=ARRAYFORMULA(IF(A' + CRAWL_FIRST_DATA_ROW + ':A="","",' +
         'IF(' + parts.positiveMatchExpr + '=0,FALSE,' +
         'IF(' + parts.excludeMatchExpr + ',FALSE,TRUE))))';
}

// Pure boolean, orthogonal to Filter Match: TRUE iff the paper is in scope
// (Filter Match=TRUE) AND a review-mode NOT group also tripped — mirrors
// jsMatchesFilterV2's `state === 'REVIEW'`. FALSE whenever Filter Match is
// FALSE too (a hard-excluded paper's review status is moot), so this column
// is never TRUE on its own without Filter Match also being TRUE.
function buildReviewFlagFormulaV2(parsedGroups, firstDetailColNum) {
  const parts = buildFilterExprPartsV2(parsedGroups, firstDetailColNum);
  if (!parts) return null;
  return '=ARRAYFORMULA(IF(A' + CRAWL_FIRST_DATA_ROW + ':A="","",' +
         'IF(' + parts.positiveMatchExpr + '=0,FALSE,' +
         'IF(' + parts.excludeMatchExpr + ',FALSE,' + parts.reviewMatchExpr + '))))';
}

// Replaces the shared applyCrawlHighlight (v1, archived) — identical term-
// helper column writing except it uses buildTermFormulaV2 (guardPhrases-
// aware) instead of the shared buildTermFormula, wires the tri-state
// formula above into Filter Match, and highlights rows by K's own text
// value directly (green=TRUE, orange=REVIEW) rather than needing a
// separate NOT-hit formula the way the pre-tri-state design did.
function applyCrawlV2Highlight(sheet, groups, guardPhrases) {
  var CRAWL_FIRST_DETAIL = CRAWL_FIRST_DETAIL_COL;

  sheet.getRange(CRAWL_HEADER_ROW, CRAWL_IN_SHEET_LINKS_COL)
    .setFormula(CRAWL_IN_SHEET_LINKS_FORMULA)
    .setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setColumnWidth(CRAWL_IN_SHEET_LINKS_COL, 90);

  var parsed = groups.map(function(g) {
    return {
      not:     g.not,
      notMode: g.notMode || 'exclude',
      terms: (g.terms || '').split(',')
        .map(function(t) { return t.trim().replace(/^["'"']|["'"']$/g, '').trim(); })
        .filter(function(t) { return t.length > 0; })
    };
  }).filter(function(g) { return g.terms.length > 0; });

  if (parsed.length === 0) return;

  var maxHelper = 60;
  sheet.getRange(1, CRAWL_FIRST_DETAIL, 1, maxHelper).breakApart().clearContent().clearFormat();
  sheet.getRange(CRAWL_HEADER_ROW, CRAWL_FIRST_DETAIL, 1, maxHelper).clearContent().clearFormat();
  sheet.getRange(CRAWL_TOTALS_ROW,  CRAWL_FIRST_DETAIL, 1, maxHelper).clearContent().clearFormat();

  var colNum = CRAWL_FIRST_DETAIL;
  parsed.forEach(function(g, groupIdx) {
    var groupStartCol = colNum;
    var isNot  = g.not;
    var isReview = isNot && g.notMode === 'review';
    var termBg = isReview ? '#fff3cd' : (isNot ? '#fce8e6' : '#e8f0fe');

    g.terms.forEach(function(term) {
      // Header (row 2): plain term text — no longer embedded in the array
      // formula (see buildTermFormulaV2), since the totals row now sits
      // between it and the real data, and an array can't spill through a
      // cell holding a separately-written formula.
      sheet.getRange(CRAWL_HEADER_ROW, colNum)
        .setValue(term)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('bottom')
        .setTextRotation(90)
        .setBackground(termBg)
        .setFontColor('#222');
      // Totals (row 3): how many rows collected so far matched this term.
      sheet.getRange(CRAWL_TOTALS_ROW, colNum)
        .setFormula(buildColumnTotalFormula(colNum))
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground(termBg)
        .setFontColor('#222');
      // Data (row 4+): the actual per-row TRUE/FALSE array, anchored one
      // row below the totals formula that reads it.
      sheet.getRange(CRAWL_FIRST_DATA_ROW, colNum)
        .setFormula(buildTermFormulaV2(term, CRAWL_COL.TITLE, CRAWL_COL.ABSTRACT, guardPhrases));
      sheet.setColumnWidth(colNum, 35);
      colNum++;
    });

    if (g.terms.length > 0) {
      var label = isReview ? 'REVIEW Group ' : (isNot ? 'NOT Group ' : 'Group ');
      label += (groupIdx + 1);
      var hdrBg = isReview ? '#f9a825' : (isNot ? '#e53935' : '#1a73e8');
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
    sheet.setRowHeight(CRAWL_HEADER_ROW, 130);

    var dataRowCount = sheet.getMaxRows() - CRAWL_FIRST_DATA_ROW + 1;
    var termDataRange = sheet.getRange(CRAWL_FIRST_DATA_ROW, CRAWL_FIRST_DETAIL, dataRowCount, colNum - CRAWL_FIRST_DETAIL);
    var startLetter = colToLetter(CRAWL_FIRST_DETAIL);
    var trueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + startLetter + CRAWL_FIRST_DATA_ROW + '=TRUE')
      .setBackground('#FFF176').setFontColor('#FFF176')
      .setRanges([termDataRange]).build();
    var falseRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=' + startLetter + CRAWL_FIRST_DATA_ROW + '=FALSE')
      .setBackground('#FFFFFF').setFontColor('#FFFFFF')
      .setRanges([termDataRange]).build();

    // Row highlight now reads K (Filter Match) and L (Review) — both pure
    // booleans — directly: green when in scope and clean, orange when in
    // scope but flagged for review. A hard-excluded row (K=FALSE) gets
    // neither, regardless of L (which the formula itself keeps FALSE
    // whenever K is FALSE, but the row rule checks both explicitly anyway
    // for clarity/defensiveness).
    var kLetter      = CRAWL_FILTER_MATCH_COL_LETTER; // "K" — from crawl.js
    var lLetter      = CRAWL_REVIEW_FLAG_COL_LETTER;  // "L" — from crawl.js
    var fullRowRange = sheet.getRange(CRAWL_FIRST_DATA_ROW, 1, dataRowCount, CRAWL_NUM_COLS);
    var trueRowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + kLetter + CRAWL_FIRST_DATA_ROW + '=TRUE,$' + lLetter + CRAWL_FIRST_DATA_ROW + '=FALSE)')
      .setBackground('#b7e1cd') // green
      .setRanges([fullRowRange]).build();
    var reviewRowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + kLetter + CRAWL_FIRST_DATA_ROW + '=TRUE,$' + lLetter + CRAWL_FIRST_DATA_ROW + '=TRUE)')
      .setBackground('#ffe0b2') // orange
      .setRanges([fullRowRange]).build();

    // Drop anything touching the term-helper columns (stale per-term rules
    // from a previous Apply Highlight Rule click) or any prior row-
    // highlight rule keyed on K or L (including the old tri-state text
    // forms ="TRUE"/="REVIEW", if this sheet predates the Match/Review
    // split, or a row-3-anchored rule from before the totals row existed)
    // — keep everything else (notably the no-abstract yellow-tint rule).
    // Matches on the column letter alone (not a specific row number) so it
    // catches a stale rule regardless of which row layout it was written
    // under.
    var existingRules = sheet.getConditionalFormatRules().filter(function(rule) {
      if (rule.getRanges().some(function(r) { return r.getColumn() >= CRAWL_FIRST_DETAIL; })) return false;
      var bc = rule.getBooleanCondition();
      if (bc) {
        var vals    = bc.getCriteriaValues();
        var formula = vals && vals[0];
        if (formula && (formula.indexOf('$' + kLetter) !== -1 || formula.indexOf('$' + lLetter) !== -1)) return false;
      }
      return true;
    });
    sheet.setConditionalFormatRules(existingRules.concat([trueRule, falseRule, trueRowRule, reviewRowRule]));
  }

  var filterFormula = buildFilterMatchFormulaV2(parsed, CRAWL_FIRST_DETAIL);
  var reviewFormula  = buildReviewFlagFormulaV2(parsed, CRAWL_FIRST_DETAIL);
  if (!filterFormula || !reviewFormula) return;
  // Header (row 2) + totals (row 3) written directly, same reasoning as the
  // term-helper columns above; the array itself (row 4+) no longer embeds
  // its own header text.
  sheet.getRange(CRAWL_HEADER_ROW, CRAWL_COL.FILTER_MATCH)
    .setValue('Filter Match')
    .setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.getRange(CRAWL_TOTALS_ROW, CRAWL_COL.FILTER_MATCH)
    .setFormula(buildColumnTotalFormula(CRAWL_COL.FILTER_MATCH))
    .setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.getRange(CRAWL_FIRST_DATA_ROW, CRAWL_COL.FILTER_MATCH)
    .setFormula(filterFormula);
  sheet.getRange(CRAWL_HEADER_ROW, CRAWL_COL.REVIEW_FLAG)
    .setValue('Review')
    .setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.getRange(CRAWL_TOTALS_ROW, CRAWL_COL.REVIEW_FLAG)
    .setFormula(buildColumnTotalFormula(CRAWL_COL.REVIEW_FLAG))
    .setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.getRange(CRAWL_FIRST_DATA_ROW, CRAWL_COL.REVIEW_FLAG)
    .setFormula(reviewFormula);
}

// Attaches a Flag Reason note to the Review cell for a REVIEW row —
// names the specific NOT group and term that triggered it, per §2 ("without
// it the reviewer has to re-derive the decision"). A cell note rather than
// a new "Flag Reason" column, consistent with this project's established
// rule against inserting columns that would shift CRAWL_COL positions for
// any crawl sheet already in progress.
function applyFlagReasonNotes(sheet, startRow, flagInfos) {
  if (!startRow) return;
  for (var i = 0; i < flagInfos.length; i++) {
    var info = flagInfos[i];
    if (!info) continue;
    var groupLabel = 'NOT group ' + (info.flagGroupIndex + 1);
    var termLabel  = info.flagTerm ? (' (term: "' + info.flagTerm + '")') : '';
    sheet.getRange(startRow + i, CRAWL_COL.REVIEW_FLAG).setNote(
      'Flagged for review: matches every positive filter group, but also ' +
      'tripped ' + groupLabel + termLabel + '. Kept and harvested, but not ' +
      'expanded — flagged rows are terminal nodes pending human triage.'
    );
  }
}

// ============================================================
// Abstract fallback — same OpenAlex recovery as crawl.js's
// fetchOpenAlexAbstract, but with a longer backoff sequence (validated
// "persist longer" decision). describeAbstractSource() from crawl.js is
// reused unmodified for the note text — its wording is generic enough for
// either pipeline.
// ============================================================

function fetchOpenAlexAbstractV2(externalIds) {
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
  for (var attempt = 0; attempt <= OPENALEX_BACKOFF_MS_V2.length; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var code = resp.getResponseCode();
      if (code === 200) {
        var data = JSON.parse(resp.getContentText());
        abstract = reconstructAbstract(data.abstract_inverted_index); // from crawl.js
        reason   = abstract ? "found" : "no-abstract";
        break;
      }
      if ((code === 429 || code >= 500) && attempt < OPENALEX_BACKOFF_MS_V2.length) {
        Utilities.sleep(OPENALEX_BACKOFF_MS_V2[attempt]);
        continue;
      }
      break;
    } catch (e) {
      if (attempt < OPENALEX_BACKOFF_MS_V2.length) {
        Utilities.sleep(OPENALEX_BACKOFF_MS_V2[attempt]);
        continue;
      }
      break;
    }
  }
  Utilities.sleep(300);
  return { abstract: abstract, reason: reason, url: pageUrl };
}

// Recovers a {MAG, DOI}-shaped externalIds from the sheet's own ID
// convention (W<mag> / S2:<hash>) for the final retry sweep, since the
// sheet doesn't otherwise persist DOI. An S2:<hash> row with no MAG carries
// nothing OpenAlex can be queried with — a known scope limit of staying
// within S2+OpenAlex rather than adding a DOI store or a third source.
function idToExternalIdsV2(id) {
  if (/^W\d+$/.test(id)) return { MAG: id.slice(1) };
  return null;
}

// ============================================================
// Phase 0 — Venue enumeration sweep (v3, off by default)
//
// "Don't name papers, name the containers" — enumerates the full
// publication record of a fixed venue list within a year window and runs
// it through the existing filter, so poster/short-paper tracks come in by
// construction (whole proceedings, not ranked search results) rather than
// needing to rank well against a topical query. Runs before Phase 1 when
// enabled; writes Direction='V' so its marginal yield is measurable
// separately from keyword-pass seeding.
//
// Resumable across trigger batches like every other phase: CRAWL2_VENUE_
// BATCH_IDX tracks which venue batch we're on, CRAWL2_VENUE_TOKEN tracks
// pagination *within* the current batch (S2's bulk-search continuation
// token — empty string means "start this batch's first page").
// ============================================================

function runVenueSweep(sheet, groups, guardPhrases, venues, yearFrom, yearTo, maxPapers, matchesOnly) {
  matchesOnly = matchesOnly !== false;
  var props = PropertiesService.getScriptProperties();
  var batchIdx  = parseInt(props.getProperty('CRAWL2_VENUE_BATCH_IDX') || '0');
  var token     = props.getProperty('CRAWL2_VENUE_TOKEN') || '';
  var collected = parseInt(props.getProperty('CRAWL2_VENUE_COUNT') || '0');
  var startTime = Date.now();

  if (!venues || venues.length === 0) {
    return { status: 'complete', message: 'Venue sweep complete — no venues configured.' };
  }

  var batches = [];
  for (var i = 0; i < venues.length; i += PHASE0_VENUE_BATCH_SIZE) {
    batches.push(venues.slice(i, i + PHASE0_VENUE_BATCH_SIZE).join(','));
  }

  var existing   = getCrawlV2ExistingKeys(sheet);
  var yearRange  = yearFrom + '-' + yearTo;

  while (batchIdx < batches.length) {
    if (Date.now() - startTime > CRAWL2_TIME_LIMIT_MS) {
      props.setProperty('CRAWL2_VENUE_BATCH_IDX', String(batchIdx));
      props.setProperty('CRAWL2_VENUE_TOKEN', token);
      props.setProperty('CRAWL2_VENUE_COUNT', String(collected));
      return { status: 'time-limit', message: 'Venue sweep — time limit reached. ' + collected +
             ' seed(s) collected so far (venue batch ' + (batchIdx + 1) + '/' + batches.length + ').' };
    }

    var result = s2BulkSearch({ venue: batches[batchIdx], year: yearRange, token: token || null });
    Utilities.sleep(1100); // same S2 pacing used elsewhere in the project

    if (result.error) {
      // Transient S2 failure even after s2Fetch's own retries — result.token
      // equals the input token here, which would otherwise look identical
      // to "batch exhausted" below and silently skip to the next venue
      // batch. Persist unchanged and ask the trigger to retry this same
      // batch/page next time instead.
      props.setProperty('CRAWL2_VENUE_BATCH_IDX', String(batchIdx));
      props.setProperty('CRAWL2_VENUE_TOKEN', token);
      props.setProperty('CRAWL2_VENUE_COUNT', String(collected));
      return { status: 'time-limit', message: 'Venue sweep — transient S2 error on batch ' + (batchIdx + 1) +
             '/' + batches.length + '; will retry. ' + collected + ' seed(s) collected so far.' };
    }

    var newRows  = [];
    var newNotes = [];
    var newFlags = []; // aligned with newRows — flag info {flagGroupIndex, flagTerm} or null
    result.data.forEach(function(paper) {
      if (!paper || !paper.paperId) return;
      var mag = paper.externalIds && paper.externalIds.MAG;
      var id  = mag ? ('W' + mag) : ('S2:' + paper.paperId);
      var doi = paper.externalIds && paper.externalIds.DOI;
      var normTitle = normalizeTitleV2(paper.title);
      if (isDuplicateCandidateV2(existing, id, normTitle, doi)) return;

      var abstract = paper.abstract || '';
      var note = null;
      if (!abstract) {
        var lookup = fetchOpenAlexAbstractV2(paper.externalIds);
        if (lookup.abstract) { abstract = lookup.abstract; paper.abstract = lookup.abstract; }
        note = describeAbstractSource(lookup); // from crawl.js
      }

      // Exhaustive by design (§3: "pass it through the existing filter") —
      // no target cap, and no separate year gate here since the venue+year
      // window is already applied server-side via the bulk-search call.
      var verdict = jsMatchesFilterV2((paper.title || '') + ' ' + abstract, groups, guardPhrases);
      if (!verdict.isMatch && matchesOnly) return; // FALSE — skip unless matchesOnly=false records it too.

      rememberCandidateV2(existing, id, normTitle, doi, verdict.isMatch);
      var row = crawlRowFromS2(paper, 0, '', 'V'); // from crawl.js — Direction='V'
      if (!verdict.expand) row[CRAWL_COL.CRAWLED - 1] = true; // REVIEW/FALSE: harvested, never expanded
      newRows.push(row);
      newNotes.push(note);
      newFlags.push(verdict.state === 'REVIEW' ? { flagGroupIndex: verdict.flagGroupIndex, flagTerm: verdict.flagTerm } : null);
      if (verdict.isMatch) collected++; // only genuine matches count toward the reported total
    });

    if (newRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd        = Math.max(0, maxPapers - currentCount);
      var writeStart     = writeCrawlRows(sheet, newRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, newNotes.slice(0, canAdd));
      applyFlagReasonNotes(sheet, writeStart, newFlags.slice(0, canAdd));
      if (canAdd < newRows.length) {
        // Cap reached — nothing more venue sweep can usefully do (unlike
        // forward, there's no "existing discovered queue" to keep draining
        // without adding rows), so finish this phase now rather than
        // stopping and requiring a manual Resume. Returning 'paper-limit'
        // here used to re-fetch and re-hit this exact same page on every
        // Resume (batchIdx/token hadn't advanced yet), repeating the
        // identical message forever instead of progressing.
        props.setProperty('CRAWL2_VENUE_BATCH_IDX', String(batchIdx));
        props.setProperty('CRAWL2_VENUE_TOKEN', '');
        props.setProperty('CRAWL2_VENUE_COUNT', String(collected));
        return { status: 'complete', message: 'Paper limit (' + maxPapers + ') reached during venue sweep — ' +
               collected + ' seed(s) collected before the cap; moving on without adding more.' };
      }
    }

    if (result.token) {
      token = result.token; // more pages remain in this venue batch
    } else {
      batchIdx++; // this batch exhausted — move to the next group of venues
      token = '';
    }
  }

  props.setProperty('CRAWL2_VENUE_BATCH_IDX', String(batchIdx));
  props.setProperty('CRAWL2_VENUE_TOKEN', '');
  props.setProperty('CRAWL2_VENUE_COUNT', String(collected));
  return { status: 'complete', message: 'Venue sweep complete. ' + collected +
         ' seed paper(s) collected across ' + venues.length + ' venue(s).' };
}

// ============================================================
// Phase 1 — Keyword pass
// ============================================================

// ============================================================
// S2 bulk-search — shared by Phase 0 (venue enumeration) and Phase 1's
// keyword pass. Unlike the relevance-ranked /paper/search (small top-N),
// /paper/search/bulk returns EVERYTHING matching the given filters, in
// pages of up to S2_BULK_PAGE_SIZE, via a continuation token — confirmed
// directly against the live endpoint before writing this (not assumed
// from memory):
//   - `query` is optional — omitting it entirely enumerates by venue/year
//     alone, which is exactly what Phase 0 needs ("the full publication
//     record", not a topical search). It also supports real boolean syntax
//     (confirmed empirically): implicit space between terms is AND-like
//     (intersection), `|` is genuine OR (union), and parenthesized groups
//     compose correctly — this is what Phase 1's split sub-queries rely on
//     (see buildKeywordSubQueries).
//   - `venue` accepts a comma-separated list and fuzzy-matches each against
//     canonical venue names (e.g. "SIGCSE" alone matched "Technical
//     Symposium on Computer Science Education").
//   - `year` takes a "YYYY-YYYY" range.
//   - `sort` accepts "publicationDate:desc".
//   - response is {total, token, data} — token is null once exhausted, and
//     `total` is populated on the very first call, before any paging.
//
// opts: { query, venue, year, sort, token, fields }. Returns
// { total, token, data } — data is [] and token is null on any HTTP error
// (treated as "nothing more this call", non-fatal per-call) rather than
// throwing, so one bad page doesn't abort an entire venue/query sweep.
function s2BulkSearch(opts) {
  var fields = opts.fields ||
    'paperId,externalIds,title,abstract,year,authors,citationCount,publicationTypes,venue';
  var params = ['fields=' + encodeURIComponent(fields)];
  if (opts.query)  params.push('query=' + encodeURIComponent(opts.query));
  if (opts.venue)  params.push('venue=' + encodeURIComponent(opts.venue));
  if (opts.year)   params.push('year=' + encodeURIComponent(opts.year));
  if (opts.sort)   params.push('sort=' + encodeURIComponent(opts.sort));
  if (opts.token)  params.push('token=' + encodeURIComponent(opts.token));

  var url  = 'https://api.semanticscholar.org/graph/v1/paper/search/bulk?' + params.join('&');
  var resp = s2Fetch(url); // from crawl.js — handles 429/5xx back-off
  // `error: true` marks a call that failed even after s2Fetch's own
  // retries — distinct from a genuine empty/exhausted result (token: null,
  // data: []) so callers can retry THIS SAME call later (e.g. via the
  // existing time-limit-and-resume path) instead of concluding "done" from
  // what's actually just a still-failing request. Preserves opts.token so
  // a retry re-fetches the same page rather than silently skipping it.
  if (resp.getResponseCode() !== 200) return { total: 0, token: opts.token || null, data: [], error: true };
  var data = JSON.parse(resp.getContentText());
  if (data.error) return { total: 0, token: opts.token || null, data: [], error: true }; // e.g. an over-broad query with no venue/year filter
  return { total: data.total || 0, token: data.token || null, data: data.data || [] };
}

// Strips common boilerplate ("Proceedings of the 2026 ACM Conference on…")
// so venue-string variants for the same conference collapse for matching
// and reporting purposes (Gap 5: "ICER appears under at least two venue
// strings"). Deliberately modest — a full canonical-venue database is out
// of scope — this only handles the common ACM/IEEE proceedings-title
// pattern, not every possible variant.
function normalizeVenueLabel(venue) {
  return (venue || '')
    .replace(/^Proceedings of the \d{4}[^-–—]*(ACM|IEEE)?\s*(Conference|Symposium|Workshop) on\s+/i, '')
    .replace(/^\d+(st|nd|rd|th)\s+/i, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function parseGroupTermsV2(group) {
  return (group.terms || '').split(',')
    .map(function(t) { return t.trim().replace(/^["']|["']$/g, '').trim(); })
    .filter(function(t) { return t.length > 0; });
}

// Formats one filter group's terms as a parenthesized OR-list — multi-word
// or hyphenated terms get quoted so S2 treats them as a single phrase
// rather than more AND'd words inside the OR list.
function buildGroupExpr(terms) {
  var alts = terms.map(function(t) {
    return (/[\s-]/.test(t)) ? ('"' + t.replace(/"/g, '') + '"') : t;
  });
  return '(' + alts.join(' | ') + ')';
}

// Splits the exhaustive keyword search into one query per term in the
// LAST positive group, each combined with the OTHER positive groups' full
// OR-lists left unsplit — rather than one single query ANDing all positive
// groups' full OR-lists together.
//
// Why: confirmed live against the API (not assumed) that S2's bulk-search
// backend is unreliable for one large compound query combining all three
// groups' full term lists (50 OR'd terms total) — it intermittently
// returns a bare "Internal Server Error" (not a rate limit), and worse,
// sometimes returns a well-formed 200 response with token: null despite
// `total` claiming many thousands more results, silently truncating the
// search with no error to catch. The identical structure with the LAST
// group's OR-list narrowed to a single term (confirmed: full Group 1 +
// full Group 2 + one Group 3 term, e.g. "LLM") paginates reliably through
// multiple pages with no such issue. Splitting on the last group
// specifically — not "whichever group has the most terms" — matters: this
// project's groups are conventionally population/modality/technology in
// that order, and only the (full G1 + full G2 + single G3 term) shape has
// actually been tested; an untested shape (e.g. single G1 term + full G2 +
// full G3) could still hit the same unreliability this exists to avoid.
//
// This trades a real increase in total API volume (a paper matching several of
// the split group's terms gets counted, and paged through, once per
// matching term — the sum of sub-query totals is larger than the single
// query's total) for actually being able to page through every sub-query
// completely, rather than fast but silently incomplete.
// Returns [{ term, query }, ...] — term is kept alongside its compound
// query string (not just the string alone) so status messages, error
// logging, and Show Crawl Progress can all say WHICH term is currently
// running/retrying, not just "sub-query 4/16".
function buildKeywordSubQueries(groups) {
  var positiveTermLists = groups
    .filter(function(g) { return !g.not; })
    .map(parseGroupTermsV2)
    .filter(function(list) { return list.length > 0; });

  if (positiveTermLists.length === 0) return [];
  if (positiveTermLists.length === 1) {
    return [{ term: '(all Group 1 terms)', query: buildGroupExpr(positiveTermLists[0]) }];
  }

  // Split on the LAST positive group specifically — not "whichever is
  // largest" (a tie between two 18-term groups would pick the first one
  // arbitrarily, and that's an untested combination). What's actually
  // confirmed reliable is full Group 1 + full Group 2 + one Group 3 term —
  // i.e. leaving the earlier (population/modality) groups fully OR'd and
  // narrowing only the last (typically "technology/tool") group, which is
  // usually the broadest, most generic-vocabulary axis anyway.
  var splitIdx = positiveTermLists.length - 1;

  var unsplitJoined = positiveTermLists
    .filter(function(_, i) { return i !== splitIdx; })
    .map(buildGroupExpr)
    .join(' ');

  return positiveTermLists[splitIdx].map(function(term) {
    return { term: term, query: unsplitJoined + ' ' + buildGroupExpr([term]) };
  });
}

// Exhaustive, deterministic keyword pass: builds the split sub-queries
// above and, for each in turn, pages through EVERY result via S2 bulk-
// search's continuation token before moving to the next — rather than
// sampling combos, and rather than one single compound query (see
// buildKeywordSubQueries for why). No target/shortfall concept any more —
// like venue sweep, it just runs to completion (or maxPapers); "how many
// matches exist" isn't a sampling question once the search is exhaustive.
// Resumable across trigger firings via CRAWL2_KEYWORD_SUBQUERY_IDX /
// _TOKEN / _PAGES_FETCHED / _COLLECTED / _RESULTS_SEEN. Every page fetch's
// own `total` field doubles as that sub-query's "N candidates found"
// preview — no separate throwaway preview call needed, unlike the old
// single-query version.
//
// opts:
//   noYearFloor:  force-ignore yearBound for this phase's own candidates
//                 even if the crawl has it configured (backward/forward
//                 keep it regardless) — opt-IN now (default false), a
//                 reversal from the old sampling-era default. When
//                 sampling combos, applying a year floor further narrowed
//                 an already-unreliable search; with exhaustive retrieval
//                 there's no such downside, so the year bound the user
//                 configured now applies here too by default — confirmed
//                 against the live API that S2 bulk-search's own `year`
//                 param genuinely narrows the server-side result set
//                 (not just a client-side post-filter), so bounding it
//                 also means fewer pages to page through, not just fewer
//                 kept rows.
//   matchesOnly:  record only genuine matches (default) vs. also record
//                 FALSE candidates for audit (per-phase toggle, unchanged).
function runKeywordPass(sheet, groups, guardPhrases, maxPapers, yearFloor, yearCeiling, yearBound, opts) {
  opts = opts || {};
  var noYearFloor = opts.noYearFloor === true;
  var matchesOnly = opts.matchesOnly !== false;
  var effectiveYearBound = noYearFloor ? false : yearBound;
  // S2's own year filter is applied server-side whenever the crawl has a
  // year bound configured — same "YYYY-YYYY" format venue sweep already
  // uses, open-ended on either side confirmed to work (e.g. "2023-" or
  // "-2020"). yearCeiling falls back to the current year, mirroring
  // isYearInBounds's own ceiling default so server-side and the client-
  // side safety-net check below can never disagree about what "in bounds"
  // means.
  var yearRangeStr = effectiveYearBound
    ? ((yearFloor || '') + '-' + (yearCeiling || new Date().getFullYear()))
    : null;

  var props = PropertiesService.getScriptProperties();
  var subQueries = buildKeywordSubQueries(groups);
  if (subQueries.length === 0) {
    return { status: 'complete', message: 'Keyword pass complete — no positive filter groups to search on. Add at least one non-NOT filter group.' };
  }

  var subQueryIdx  = parseInt(props.getProperty('CRAWL2_KEYWORD_SUBQUERY_IDX') || '0');
  var token        = props.getProperty('CRAWL2_KEYWORD_TOKEN') || null;
  var pagesFetched = parseInt(props.getProperty('CRAWL2_KEYWORD_PAGES_FETCHED') || '0');
  var collected    = parseInt(props.getProperty('CRAWL2_KEYWORD_COLLECTED')     || '0');
  var resultsSeen  = parseInt(props.getProperty('CRAWL2_KEYWORD_RESULTS_SEEN')  || '0');
  // Diagnostics for "which sub-query is actually flaky" — subQueryErrors
  // resets to 0 whenever subQueryIdx advances (it's specific to whichever
  // sub-query is currently running); totalErrors accumulates across the
  // whole keyword pass; retriedTerms records each sub-query's term ONCE,
  // the first time it hits an error, so the final summary can name exactly
  // which terms needed a retry without needing to re-derive it later.
  var subQueryErrors = parseInt(props.getProperty('CRAWL2_KEYWORD_SUBQUERY_ERRORS') || '0');
  var totalErrors    = parseInt(props.getProperty('CRAWL2_KEYWORD_TOTAL_ERRORS')    || '0');
  var retriedTerms   = JSON.parse(props.getProperty('CRAWL2_KEYWORD_RETRIED_TERMS') || '[]');
  var startTime = Date.now();

  var existing = getCrawlV2ExistingKeys(sheet);

  function persistProgress() {
    props.setProperty('CRAWL2_KEYWORD_SUBQUERY_IDX', String(subQueryIdx));
    props.setProperty('CRAWL2_KEYWORD_TOKEN', token || '');
    props.setProperty('CRAWL2_KEYWORD_PAGES_FETCHED', String(pagesFetched));
    props.setProperty('CRAWL2_KEYWORD_COLLECTED', String(collected));
    props.setProperty('CRAWL2_KEYWORD_RESULTS_SEEN', String(resultsSeen));
    props.setProperty('CRAWL2_KEYWORD_SUBQUERY_ERRORS', String(subQueryErrors));
    props.setProperty('CRAWL2_KEYWORD_TOTAL_ERRORS', String(totalErrors));
    props.setProperty('CRAWL2_KEYWORD_RETRIED_TERMS', JSON.stringify(retriedTerms));
  }

  while (subQueryIdx < subQueries.length) {
    var currentTerm = subQueries[subQueryIdx].term;

    if (Date.now() - startTime > CRAWL2_TIME_LIMIT_MS) {
      persistProgress();
      return { status: 'time-limit', message: 'Keyword pass — time limit reached. Sub-query ' + (subQueryIdx + 1) +
             '/' + subQueries.length + ' ("' + currentTerm + '"), ' + pagesFetched + ' page(s) fetched so far, ' +
             collected + ' match(es) collected (' + resultsSeen + ' candidates examined).' };
    }

    var pageResult = s2BulkSearch({ query: subQueries[subQueryIdx].query, year: yearRangeStr, token: token || null });

    if (pageResult.error) {
      // Transient S2 failure even after s2Fetch's own retries — persist
      // unchanged (token echoes back the input token, so this exact page
      // of this exact sub-query gets retried) and ask the trigger to
      // re-fire, rather than falling through to the token-based
      // completion check below (which would otherwise treat the echoed-
      // back token identically to "this sub-query is exhausted").
      subQueryErrors++;
      totalErrors++;
      if (subQueryErrors === 1 && retriedTerms.indexOf(currentTerm) === -1) retriedTerms.push(currentTerm);
      persistProgress();
      return { status: 'time-limit', message: 'Keyword pass — transient S2 error on sub-query ' + (subQueryIdx + 1) +
             '/' + subQueries.length + ' ("' + currentTerm + '", attempt ' + subQueryErrors + ' on this sub-query, ' +
             totalErrors + ' total this pass); will retry. ' + collected + ' match(es) collected so far (' +
             resultsSeen + ' candidates examined).' };
    }
    Utilities.sleep(1100); // same S2 pacing used elsewhere in the project

    if (token == null) {
      // First page of this sub-query — its own `total` field doubles as
      // the "N candidates found" preview, no separate call needed.
      setCrawlStatus(sheet, 'Keyword pass — sub-query ' + (subQueryIdx + 1) + '/' + subQueries.length +
        ' ("' + currentTerm + '", ' + (pageResult.total || 0) + ' candidate(s)), retrieving details…');
    }

    pagesFetched++;
    resultsSeen += pageResult.data.length;

    var newRows  = [];
    var newNotes = [];
    var newFlags = []; // aligned with newRows — flag info {flagGroupIndex, flagTerm} or null
    pageResult.data.forEach(function(paper) {
      if (!paper || !paper.paperId) return;
      var mag = paper.externalIds && paper.externalIds.MAG;
      var id  = mag ? ('W' + mag) : ('S2:' + paper.paperId);
      var doi = paper.externalIds && paper.externalIds.DOI;
      var normTitle = normalizeTitleV2(paper.title);
      if (isDuplicateCandidateV2(existing, id, normTitle, doi)) return;

      var abstract = paper.abstract || '';
      var note = null;
      if (!abstract) {
        var lookup = fetchOpenAlexAbstractV2(paper.externalIds);
        if (lookup.abstract) { abstract = lookup.abstract; paper.abstract = lookup.abstract; }
        note = describeAbstractSource(lookup); // from crawl.js
      }

      var yearOk  = effectiveYearBound ? isYearInBounds(paper.year, yearFloor, yearCeiling, effectiveYearBound) : true; // from crawl.js
      var verdict = jsMatchesFilterV2((paper.title || '') + ' ' + abstract, groups, guardPhrases);
      var keep    = verdict.isMatch && yearOk;   // TRUE or REVIEW, and within the year bound
      var expand  = verdict.expand  && yearOk;   // TRUE only — REVIEW is always terminal
      if (!keep && matchesOnly) return; // FALSE, or year-rejected — skip unless matchesOnly=false records it too.

      rememberCandidateV2(existing, id, normTitle, doi, keep);
      var row = crawlRowFromS2(paper, 0, '', 'K'); // from crawl.js
      if (!expand) row[CRAWL_COL.CRAWLED - 1] = true; // REVIEW/FALSE/year-rejected: harvested, never expanded
      newRows.push(row);
      newNotes.push(note);
      // Only a genuine REVIEW verdict carries the flag note — see forward
      // pass's identical comment for why yearOk-driven exclusions don't.
      newFlags.push((keep && verdict.state === 'REVIEW') ? { flagGroupIndex: verdict.flagGroupIndex, flagTerm: verdict.flagTerm } : null);
      if (keep) collected++;
    });

    if (newRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd        = Math.max(0, maxPapers - currentCount);
      var writeStart     = writeCrawlRows(sheet, newRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, newNotes.slice(0, canAdd));
      applyFlagReasonNotes(sheet, writeStart, newFlags.slice(0, canAdd));
      if (canAdd < newRows.length) {
        // Cap reached — same reasoning as venue sweep: nothing more useful
        // to do (every remaining page can only ever be discarded), so
        // finish now rather than keep paging just to throw results away.
        persistProgress();
        return { status: 'complete', message: 'Paper limit (' + maxPapers + ') reached during keyword pass — ' +
               collected + ' match(es) collected before the cap (' + resultsSeen +
               ' candidates examined); moving on without adding more.' };
      }
    }

    token = pageResult.token;
    if (!token) {
      // This sub-query exhausted — advance to the next one, if any. Reset
      // the per-sub-query error counter (it's specific to whichever
      // sub-query is currently running); totalErrors/retriedTerms are
      // cumulative for the whole pass and don't reset here.
      subQueryIdx++;
      token = null;
      subQueryErrors = 0;
    }
  }

  persistProgress();
  var summary = 'Keyword pass complete. ' + collected + ' match(es) collected from ' +
    pagesFetched + ' page' + (pagesFetched === 1 ? '' : 's') + ' across ' + subQueries.length + ' sub-quer' +
    (subQueries.length === 1 ? 'y' : 'ies') + ' (' + resultsSeen + ' candidates examined).';
  if (retriedTerms.length > 0) {
    summary += ' ' + retriedTerms.length + ' sub-quer' + (retriedTerms.length === 1 ? 'y' : 'ies') +
      ' needed at least one retry due to transient S2 errors (' + totalErrors + ' total): ' +
      retriedTerms.join(', ') + '.';
  }
  return { status: 'complete', message: summary };
}

// ============================================================
// Phase 2 — Backward pass
//
// Per the brief's §2/§3 (validated, unchanged from the brief): only papers
// with Filter Match=TRUE expand backward, using the SAME depth gate as
// forward (not v1's separate, smaller backwardDepth setting — the v2
// pipeline runs backward over keyword-pass seeds directly rather than deep
// in a forward-discovered tree, so that independent cap no longer applies
// the way it did in v1).
//
// "Filter Match=TRUE" is recomputed here from each row's own stored
// Title+Abstract via jsMatchesFilterV2, rather than read from the sheet's
// live formula column — the live formula only reflects reality once Apply
// Highlight Rule has been run, so re-deriving it in-memory is the more
// robust source of truth for phase-gating.
//
// Already scans the WHOLE sheet regardless of Direction (K/B/F all
// eligible as long as Filter Match=TRUE and depth < maxDepth) — this is
// also what makes a "second pass after forward" (v3 §4/§7.2, sign-off:
// sequential, not interleaved) trivial to add: the SAME function, called
// again with a different propPrefix so its own resumability state doesn't
// collide with the first pass, naturally picks up every F-discovered match
// forward produced, on top of whatever it already covered. It also
// re-examines K/B parents already processed in pass 1 (a real but accepted
// inefficiency — avoiding that would need a persistent "already a backward
// source" ID set threaded across passes, more complexity than the brief's
// "add a second pass" ask calls for).
//
// propPrefix namespaces this pass's own resumability + instrumentation
// properties — 'CRAWL2_BACKWARD' for pass 1 (unchanged), 'CRAWL2_BACKWARD2'
// for the optional second pass.
function runBackwardPassV2(sheet, groups, guardPhrases, maxDepth, maxPapers, yearFloor, yearCeiling, yearBound, propPrefix, matchesOnly) {
  propPrefix = propPrefix || 'CRAWL2_BACKWARD';
  matchesOnly = matchesOnly !== false;
  var startTime = Date.now();
  var props     = PropertiesService.getScriptProperties();
  var idx       = parseInt(props.getProperty(propPrefix + '_IDX') || '0');
  // Reject-counting (v3 §4: "a per-parent count of references examined vs
  // kept, so the phase has a measurable hit rate") — aggregate, not
  // per-parent-logged, per the brief's own "at minimum" allowance.
  var examined  = parseInt(props.getProperty(propPrefix + '_EXAMINED') || '0');
  var kept      = parseInt(props.getProperty(propPrefix + '_KEPT')     || '0');

  var lastRow = getCrawlLastDataRow(sheet);
  if (lastRow < CRAWL_FIRST_DATA_ROW) return { status: 'complete', message: 'Backward pass complete. No papers found.' };

  var numRows = lastRow - CRAWL_FIRST_DATA_ROW + 1;
  var data    = sheet.getRange(CRAWL_FIRST_DATA_ROW, 1, numRows, CRAWL_NUM_COLS).getValues();

  var existing        = getCrawlV2ExistingKeys(sheet);
  var paperIds         = [];
  var paperSheetRows   = [];
  var paperDepths      = [];
  data.forEach(function(row, i) {
    var id = String(row[CRAWL_COL.ID - 1] || '').trim();
    if (!id) return;
    var depth = parseInt(row[CRAWL_COL.DEPTH - 1]) || 0;
    if (depth >= maxDepth) return; // harvest-not-expand — same gate as forward
    var title    = String(row[CRAWL_COL.TITLE - 1]    || '');
    var abstract = String(row[CRAWL_COL.ABSTRACT - 1] || '');
    var verdict  = jsMatchesFilterV2(title + ' ' + abstract, groups, guardPhrases);
    if (!verdict.expand) return; // only Filter Match=TRUE papers expand backward — REVIEW rows are terminal
    paperIds.push(id);
    paperSheetRows.push(CRAWL_FIRST_DATA_ROW + i);
    paperDepths.push(depth);
  });

  function persistProgress() {
    props.setProperty(propPrefix + '_IDX', String(idx));
    props.setProperty(propPrefix + '_EXAMINED', String(examined));
    props.setProperty(propPrefix + '_KEPT', String(kept));
  }

  var processed = 0;
  while (idx < paperIds.length) {
    if (Date.now() - startTime > CRAWL2_TIME_LIMIT_MS) {
      persistProgress();
      updateCrawlMatchedCites(sheet); // from crawl.js
      return { status: 'time-limit', message: 'Backward pass — time limit reached. Processed ' + processed +
             '; ' + (paperIds.length - idx) + ' remain.' };
    }

    var paperSheetRow = paperSheetRows[idx];
    var paperDepth     = paperDepths[idx];
    var paperId        = paperIds[idx++];
    processed++;

    Utilities.sleep(1100);
    var refs = [];
    try { refs = s2GetReferences(paperId); } catch (e) { // from crawl.js
      markFetchFailure(sheet, paperSheetRow, e); // from crawl.js
      continue;
    }
    examined += refs.length;

    var newRows  = [];
    var newNotes = [];
    var newFlags = []; // aligned with newRows — flag info {flagGroupIndex, flagTerm} or null
    refs.forEach(function(ref) {
      if (!ref || !ref.paperId) return;
      var mag   = ref.externalIds && ref.externalIds.MAG;
      var refId = mag ? ('W' + mag) : ('S2:' + ref.paperId);
      var doi   = ref.externalIds && ref.externalIds.DOI;
      var normTitle = normalizeTitleV2(ref.title);
      if (isDuplicateCandidateV2(existing, refId, normTitle, doi)) return;

      var note = null;
      if (!ref.abstract) {
        var lookup = fetchOpenAlexAbstractV2(ref.externalIds);
        if (lookup.abstract) ref.abstract = lookup.abstract;
        note = describeAbstractSource(lookup);
      }

      var yearOk  = isYearInBounds(ref.year, yearFloor, yearCeiling, yearBound);
      var verdict = jsMatchesFilterV2((ref.title || '') + ' ' + (ref.abstract || ''), groups, guardPhrases);
      var keep    = verdict.isMatch && yearOk;   // TRUE or REVIEW, and within the year bound
      var expand  = verdict.expand  && yearOk;   // TRUE only — REVIEW is always terminal
      if (!keep && matchesOnly) return; // FALSE, or year-rejected — skip unless matchesOnly=false records it too.

      rememberCandidateV2(existing, refId, normTitle, doi, keep);
      var row = crawlRowFromS2(ref, paperDepth + 1, paperId, 'B'); // from crawl.js
      if (!expand) row[CRAWL_COL.CRAWLED - 1] = true; // REVIEW/FALSE/year-rejected: harvested, never expanded
      newRows.push(row);
      newNotes.push(note);
      // Only a genuine REVIEW verdict carries the flag note — see forward
      // pass's identical comment for why yearOk-driven exclusions don't.
      newFlags.push((keep && verdict.state === 'REVIEW') ? { flagGroupIndex: verdict.flagGroupIndex, flagTerm: verdict.flagTerm } : null);
      if (keep) kept++; // hit-rate stat stays tied to genuine, in-range matches, not rows written
    });

    if (newRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd        = Math.max(0, maxPapers - currentCount);
      var writeStart     = writeCrawlRows(sheet, newRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, newNotes.slice(0, canAdd));
      applyFlagReasonNotes(sheet, writeStart, newFlags.slice(0, canAdd));
    }
  }

  persistProgress();
  updateCrawlMatchedCites(sheet);
  return { status: 'complete', message: 'Backward pass complete. Processed ' + processed + ' matching paper(s), ' +
         'examined ' + examined + ' reference(s), kept ' + kept +
         ' (' + (examined > 0 ? Math.round(100 * kept / examined) : 0) + '% hit rate).' };
}

// ============================================================
// Phase 3 — Forward pass
//
// Mirrors crawl.js's runCrawlLoop, but uses the score-demotion filter
// (jsMatchesFilterV2) instead of the hard-veto jsMatchesFilter, and the
// title+ID dedup instead of ID-only — the two validated behavior changes
// for v2. Otherwise identical: same time-limit batching, same paper-cap
// handling, same depth harvest-not-expand gate, same OpenAlex abstract
// fallback pattern (with the longer v2 backoff).
// ============================================================

function runForwardPassV2(sheet, groups, guardPhrases, maxDepth, maxPapers, matchesOnly, yearFloor, yearCeiling, yearBound) {
  matchesOnly = matchesOnly !== false;
  var startTime = Date.now();
  var processed  = 0;
  // Rows that failed this session and are being left queued for a later
  // retry (markFetchFailure returned true) — skipped for the REST of this
  // pass so the loop doesn't immediately re-select the same failing row
  // over and over and burn the whole time budget on it; a future session
  // starts this set fresh and retries normally.
  var retriedThisPass = new Set();

  while (true) {
    if (Date.now() - startTime > CRAWL2_TIME_LIMIT_MS) {
      updateCrawlMatchedCites(sheet);
      var remaining = countUncrawled(sheet); // from crawl.js
      return { status: 'time-limit', message: "Forward pass — time limit reached. Processed " + processed +
             " papers this session; " + remaining + " remain in queue. Click Resume Crawl v2 to continue." };
    }

    var next = findNextUncrawled(sheet, retriedThisPass); // from crawl.js
    if (!next) {
      updateCrawlMatchedCites(sheet);
      return { status: 'complete', message: "Forward pass complete. " + processed + " papers processed in this session." };
    }

    var sheetRow = next.sheetRow;
    var depth    = next.depth;
    var id       = next.id;
    var title    = next.title;

    if (depth >= maxDepth) {
      sheet.getRange(sheetRow, CRAWL_COL.CRAWLED).setValue(true);
      processed++;
      continue;
    }

    // Already at/over the cap (set once below, the first time this happens)
    // — keep draining the rest of the already-discovered queue (mark
    // Crawled=true) without fetching candidates for it at all, since
    // nothing found could ever be written anyway. Skipping the fetch here
    // (rather than fetching and then discarding at the write step) avoids
    // burning an S2 call — and an API rate-limit backoff — per remaining
    // queued paper for no benefit. This is what lets the crawl finish
    // unattended once the cap is hit, instead of requiring a fresh manual
    // Resume for every single subsequent paper in the queue.
    if (getCrawlLastDataRow(sheet) - 2 >= maxPapers) {
      sheet.getRange(sheetRow, CRAWL_COL.CRAWLED).setValue(true);
      processed++;
      continue;
    }

    var candidates = [];
    try {
      candidates = fetchForwardCandidates(id, title); // from crawl.js
    } catch (e) {
      var shouldRetry = markFetchFailure(sheet, sheetRow, e); // from crawl.js
      if (shouldRetry) { retriedThisPass.add(sheetRow); } else { sheet.getRange(sheetRow, CRAWL_COL.CRAWLED).setValue(true); }
      processed++;
      continue;
    }

    Utilities.sleep(1100);

    var existing      = getCrawlV2ExistingKeys(sheet);
    var matchRows     = [];
    var abstractNotes = [];
    var flagInfos     = []; // aligned with matchRows — flag info {flagGroupIndex, flagTerm} or null
    for (var i = 0; i < candidates.length; i++) {
      var c   = candidates[i];
      var mag = c.externalIds && c.externalIds.MAG;
      var cId = mag ? ('W' + mag) : ('S2:' + c.paperId);
      var cDoi = c.externalIds && c.externalIds.DOI;
      var normTitle = normalizeTitleV2(c.title);
      if (isDuplicateCandidateV2(existing, cId, normTitle, cDoi)) continue;

      var abstract = c.abstract || '';
      var note = null;
      if (!abstract) {
        var lookup = fetchOpenAlexAbstractV2(c.externalIds);
        if (lookup.abstract) { abstract = lookup.abstract; c.abstract = lookup.abstract; }
        note = describeAbstractSource(lookup);
      }

      var yearOk  = isYearInBounds(getCandidateYear(c, "forward"), yearFloor, yearCeiling, yearBound); // from crawl.js
      var verdict = jsMatchesFilterV2((c.title || '') + ' ' + abstract, groups, guardPhrases);
      var keep    = verdict.isMatch && yearOk;   // TRUE or REVIEW, and within the year bound
      var expand  = verdict.expand  && yearOk;   // TRUE only — REVIEW is always terminal
      if (!keep && matchesOnly) continue;
      var row = crawlRowFromS2(c, depth + 1, id, 'F'); // from crawl.js
      if (!expand) row[CRAWL_COL.CRAWLED - 1] = true;
      matchRows.push(row);
      abstractNotes.push(note);
      // Only a genuine REVIEW verdict carries the flag note — if yearOk is
      // what actually excluded it, a "flagged for review" note would be a
      // misleading explanation for why the row is there.
      flagInfos.push((keep && verdict.state === 'REVIEW') ? { flagGroupIndex: verdict.flagGroupIndex, flagTerm: verdict.flagTerm } : null);
      rememberCandidateV2(existing, cId, normTitle, cDoi, keep);
    }

    if (matchRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd        = Math.max(0, maxPapers - currentCount);
      var writeStart     = writeCrawlRows(sheet, matchRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, abstractNotes.slice(0, canAdd));
      applyFlagReasonNotes(sheet, writeStart, flagInfos.slice(0, canAdd));
      if (canAdd < matchRows.length) {
        // First time crossing maxPapers — surface it once (status + log),
        // then keep going rather than stopping: every iteration from here
        // on takes the "already at/over the cap" branch above, so the rest
        // of the existing queue still gets marked Crawled=true (and its own
        // forward candidates discarded) automatically, with no further
        // rows added and no manual Resume needed for the crawl to reach
        // Complete. Doesn't return, doesn't delete the trigger — this is a
        // one-time notification, not a stop condition any more.
        updateCrawlMatchedCites(sheet);
        var logRowPL = parseInt(PropertiesService.getScriptProperties().getProperty('CRAWL2_LOG_ROW') || '0') || 0;
        updateLogRow(logRowPL, 'Paper Limit');
        setCrawlStatus(sheet, 'Paper limit (' + maxPapers + ') reached — no further papers will be added, but the ' +
          'crawl is continuing to process the rest of the existing queue automatically.');
      }
    }

    sheet.getRange(sheetRow, CRAWL_COL.CRAWLED).setValue(true);
    processed++;
  }
}

// ============================================================
// Final phase — abstract retry sweep
//
// Validated response to "persist with the search for longer before moving
// on": one more attempt at every row still missing an abstract, after all
// three phases finish, before the run is marked Complete. Resumable via
// CRAWL2_SWEEP_IDX like every other phase, since a large sheet's worth of
// missing abstracts could itself exceed the 6-minute execution limit.
// ============================================================

function runAbstractRetrySweep(sheet, startIdx) {
  var startTime = Date.now();
  var lastRow   = getCrawlLastDataRow(sheet);
  if (lastRow < CRAWL_FIRST_DATA_ROW) return { status: 'complete', recovered: 0, nextIdx: 0 };
  var numRows = lastRow - CRAWL_FIRST_DATA_ROW + 1;
  var idCol   = sheet.getRange(CRAWL_FIRST_DATA_ROW, CRAWL_COL.ID,       numRows, 1).getValues();
  var absCol  = sheet.getRange(CRAWL_FIRST_DATA_ROW, CRAWL_COL.ABSTRACT, numRows, 1).getValues();
  var recovered = 0;
  var i = startIdx || 0;

  for (; i < numRows; i++) {
    if (Date.now() - startTime > CRAWL2_TIME_LIMIT_MS) {
      return { status: 'time-limit', recovered: recovered, nextIdx: i };
    }
    var abstract = String(absCol[i][0] || '').trim();
    if (abstract) continue;
    var id = String(idCol[i][0] || '').trim();
    if (!id) continue;
    // S2:<hash>-only rows carry no MAG/DOI to check OpenAlex with — a known
    // scope limit of staying within S2+OpenAlex (no persisted DOI store).
    var externalIds = idToExternalIdsV2(id);
    if (!externalIds) continue;

    var lookup = fetchOpenAlexAbstractV2(externalIds);
    var row = CRAWL_FIRST_DATA_ROW + i;
    if (lookup.abstract) {
      sheet.getRange(row, CRAWL_COL.ABSTRACT).setValue(lookup.abstract)
        .setNote(describeAbstractSource(lookup) + ' (recovered on final retry sweep)');
      recovered++;
    } else {
      sheet.getRange(row, CRAWL_COL.ABSTRACT)
        .setNote(describeAbstractSource(lookup) + ' (rechecked on final retry sweep, still unresolved)');
    }
    Utilities.sleep(300);
  }
  return { status: 'complete', recovered: recovered, nextIdx: i };
}

// ============================================================
// Background batch handler — the v2 trigger entry point.
// Phase order (v22 §0.1/§1: promoted to default, no longer gated by
// on/off flags): venue → keyword → backward → forward → backward2 (second
// backward pass, over the now-larger match set forward produced) → sweep
// → complete. Venue is a no-op if no venues are configured; every other
// phase always runs. matchesOnly/yearBound/maxDepth/maxPapers/backward
// depth remain per-run configuration, just no longer an on/off switch for
// whole phases.
// ============================================================

function crawlV2BatchTrigger() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return;

  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('CRAWL2_CONSEC_FAILURES', '0');
    var sheetName = props.getProperty('CRAWL2_ACTIVE_SHEET');
    if (!sheetName) { deleteCrawlV2Trigger(); return; }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) { deleteCrawlV2Trigger(); return; }

    var phase       = props.getProperty('CRAWL2_PHASE') || 'venue';
    updateCrawlV2Timing(sheet, phase); // refresh Duration + Progress on every firing, before this batch's own work
    var batch       = parseInt(props.getProperty('CRAWL2_BATCH_NUM') || '1');
    var groups      = JSON.parse(props.getProperty('CRAWL2_FILTER_GROUPS') || '[]');
    var guardPhrases = JSON.parse(props.getProperty('CRAWL2_GUARD_PHRASES') || '[]');
    var maxDepth    = parseInt(props.getProperty('CRAWL2_MAX_DEPTH')  || '2');
    // Backward's own depth, decoupled from forward's — backward chases
    // REFERENCES (older work), which compounds fast into ancient, generic
    // citation noise (statistics-method papers, foundational psychology,
    // etc. cited alongside the real topical references) the more
    // generations it's allowed to recurse. Defaults to maxDepth (the old,
    // shared-depth behaviour) for crawls started before this existed.
    var backwardMaxDepth = parseInt(props.getProperty('CRAWL2_BACKWARD_MAX_DEPTH') || String(maxDepth));
    var maxPapers   = parseInt(props.getProperty('CRAWL2_MAX_PAPERS') || '300');
    // Per-phase (§ "log non-matches by phase" — venue sweep in particular
    // can produce a lot of noise since it's exhaustive by design, so it's
    // useful to record non-matches for keyword/backward/forward without
    // also flooding the sheet with every non-matching venue paper).
    // CRAWL2_MATCHES_ONLY (no suffix) is the pre-per-phase property name —
    // read as a fallback default for crawls started before this change.
    var legacyMatchesOnly = props.getProperty('CRAWL2_MATCHES_ONLY') !== 'false';
    var matchesOnlyVenue    = props.getProperty('CRAWL2_MATCHES_ONLY_VENUE')    != null ? props.getProperty('CRAWL2_MATCHES_ONLY_VENUE')    !== 'false' : legacyMatchesOnly;
    var matchesOnlyKeyword  = props.getProperty('CRAWL2_MATCHES_ONLY_KEYWORD')  != null ? props.getProperty('CRAWL2_MATCHES_ONLY_KEYWORD')  !== 'false' : legacyMatchesOnly;
    var matchesOnlyBackward = props.getProperty('CRAWL2_MATCHES_ONLY_BACKWARD') != null ? props.getProperty('CRAWL2_MATCHES_ONLY_BACKWARD') !== 'false' : legacyMatchesOnly;
    var matchesOnlyForward  = props.getProperty('CRAWL2_MATCHES_ONLY_FORWARD')  != null ? props.getProperty('CRAWL2_MATCHES_ONLY_FORWARD')  !== 'false' : legacyMatchesOnly;
    var yearFloor   = parseInt(props.getProperty('CRAWL2_YEAR_FLOOR')   || '0') || 0;
    var yearCeiling = parseInt(props.getProperty('CRAWL2_YEAR_CEILING') || '0') || 0;
    var yearBound   = props.getProperty('CRAWL2_YEAR_BOUND') !== 'false';
    var logRow      = parseInt(props.getProperty('CRAWL2_LOG_ROW') || '0') || 0;

    // Per-phase on/off toggles — default true (today's full pipeline) for
    // crawls started before these existed. 'backward'/'forward'/'backward2'
    // are only ever entered via a transition below that already checked the
    // relevant flag, so those phase blocks themselves don't need to
    // re-check — only 'venue' (always the hardcoded starting phase) and the
    // transition logic (nextPhaseAfter) need to know about these.
    var runVenuePhase    = props.getProperty('CRAWL2_RUN_VENUE_PHASE')    !== 'false';
    var runBackwardPhase = props.getProperty('CRAWL2_RUN_BACKWARD_PHASE') !== 'false';
    var runForwardPhase  = props.getProperty('CRAWL2_RUN_FORWARD_PHASE')  !== 'false';

    // Returns the next phase to run after `current` completes, skipping any
    // disabled ones — lets a keyword-only (or keyword + one direction) run
    // work without restructuring the underlying venue -> keyword ->
    // backward -> forward -> backward2 -> sweep sequence itself. Backward2
    // is skipped whenever backward pass 1 was skipped too — it exists
    // specifically to cover what forward finds on top of backward's own
    // sources, so it's redundant with pass 1 if that never ran.
    function nextPhaseAfter(current) {
      if (current === 'venue')     return 'keyword';
      if (current === 'keyword')   return runBackwardPhase ? 'backward' : (runForwardPhase ? 'forward' : 'sweep');
      if (current === 'backward')  return runForwardPhase ? 'forward' : 'sweep';
      if (current === 'forward')   return runBackwardPhase ? 'backward2' : 'sweep';
      return 'sweep';
    }

    var phase0Venues   = JSON.parse(props.getProperty('CRAWL2_PHASE0_VENUES') || '[]');
    var phase0YearFrom = parseInt(props.getProperty('CRAWL2_PHASE0_YEAR_FROM') || String(PHASE0_YEAR_FROM_DEFAULT));
    var phase0YearTo   = parseInt(props.getProperty('CRAWL2_PHASE0_YEAR_TO')   || String(PHASE0_YEAR_TO_DEFAULT));
    var keywordOpts = {
      matchesOnly: matchesOnlyKeyword
      // noYearFloor deliberately omitted — runKeywordPass now respects
      // the crawl's own yearBound/yearFloor/yearCeiling by default (opt
      // IN to ignore it via noYearFloor:true, not opt out).
    };

    var result;

    if (phase === 'venue') {
      if (!runVenuePhase) {
        var nextP = nextPhaseAfter('venue');
        props.setProperty('CRAWL2_PHASE', nextP);
        props.setProperty('CRAWL2_BATCH_NUM', '1');
        setCrawlStatus(sheet, 'Venue sweep disabled — starting keyword pass…');
      } else {
        setCrawlStatus(sheet, 'Venue sweep — batch ' + batch + '…');
        result = runVenueSweep(sheet, groups, guardPhrases, phase0Venues, phase0YearFrom, phase0YearTo, maxPapers, matchesOnlyVenue);
        if (result.status === 'time-limit') {
          props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
          setCrawlStatus(sheet, result.message);
        } else {
          // 'complete' covers both a genuinely finished sweep and the
          // paper-limit-reached case (runVenueSweep returns 'complete' for
          // both now — nothing more for this phase to do once the cap is
          // hit, same as a normal finish) — either way, move on to keyword.
          props.setProperty('CRAWL2_PHASE', 'keyword');
          props.setProperty('CRAWL2_BATCH_NUM', '1');
          setCrawlStatus(sheet, result.message + ' — starting keyword pass…');
        }
      }

    } else if (phase === 'keyword') {
      setCrawlStatus(sheet, 'Keyword pass — batch ' + batch + '…');
      result = runKeywordPass(sheet, groups, guardPhrases, maxPapers, yearFloor, yearCeiling, yearBound, keywordOpts);
      if (result.status === 'time-limit') {
        props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, result.message);
      } else {
        // 'complete' covers both a genuine finish and the paper-limit-
        // reached case (runKeywordPass returns 'complete' for both now,
        // same reasoning as venue sweep) — either way, move on to whatever
        // phase is next enabled (backward/forward/sweep — see nextPhaseAfter).
        var nextP = nextPhaseAfter('keyword');
        props.setProperty('CRAWL2_PHASE', nextP);
        props.setProperty('CRAWL2_BATCH_NUM', '1');
        setCrawlStatus(sheet, result.message + ' — starting ' + nextP + ' phase…');
      }

    } else if (phase === 'backward') {
      setCrawlStatus(sheet, 'Backward pass — batch ' + batch + '…');
      result = runBackwardPassV2(sheet, groups, guardPhrases, backwardMaxDepth, maxPapers, yearFloor, yearCeiling, yearBound, 'CRAWL2_BACKWARD', matchesOnlyBackward);
      if (result.status === 'time-limit') {
        props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, result.message);
      } else {
        var nextP = nextPhaseAfter('backward');
        props.setProperty('CRAWL2_PHASE', nextP);
        props.setProperty('CRAWL2_BATCH_NUM', '1');
        setCrawlStatus(sheet, result.message + ' — starting ' + nextP + ' phase…');
      }

    } else if (phase === 'forward') {
      setCrawlStatus(sheet, 'Forward pass — batch ' + batch + '…');
      result = runForwardPassV2(sheet, groups, guardPhrases, maxDepth, maxPapers, matchesOnlyForward, yearFloor, yearCeiling, yearBound);
      if (result.status === 'time-limit') {
        props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, result.message);
      } else {
        // 'complete' covers both a genuine finish and having hit maxPapers
        // partway through (runForwardPassV2 now drains the rest of the
        // queue itself once the cap is reached, rather than stopping).
        var nextP = nextPhaseAfter('forward');
        props.setProperty('CRAWL2_PHASE', nextP);
        props.setProperty('CRAWL2_BATCH_NUM', '1');
        setCrawlStatus(sheet, result.message + ' — starting ' + nextP + ' phase…');
      }

    } else if (phase === 'backward2') {
      setCrawlStatus(sheet, 'Second backward pass — batch ' + batch + '…');
      result = runBackwardPassV2(sheet, groups, guardPhrases, backwardMaxDepth, maxPapers, yearFloor, yearCeiling, yearBound, 'CRAWL2_BACKWARD2', matchesOnlyBackward);
      if (result.status === 'time-limit') {
        props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, result.message);
      } else {
        props.setProperty('CRAWL2_PHASE', 'sweep');
        setCrawlStatus(sheet, result.message + ' — running final abstract retry sweep…');
      }

    } else if (phase === 'sweep') {
      var sweepIdx   = parseInt(props.getProperty('CRAWL2_SWEEP_IDX')       || '0');
      var sweepTotal = parseInt(props.getProperty('CRAWL2_SWEEP_RECOVERED') || '0');
      var sweepResult = runAbstractRetrySweep(sheet, sweepIdx);
      sweepTotal += sweepResult.recovered;
      if (sweepResult.status === 'time-limit') {
        props.setProperty('CRAWL2_SWEEP_IDX', String(sweepResult.nextIdx));
        props.setProperty('CRAWL2_SWEEP_RECOVERED', String(sweepTotal));
        setCrawlStatus(sheet, 'Final abstract retry sweep — ' + sweepTotal + ' recovered so far…');
      } else {
        deleteCrawlV2Trigger();
        updateLogRow(logRow, 'Complete');
        setCrawlStatus(sheet, 'Complete (' + sweepTotal + ' abstract(s) recovered on final sweep)');
        updateCrawlV2Timing(sheet, 'complete'); // final Duration freeze + Progress = "Complete"
      }
    }

  } catch (e) {
    var propsE    = PropertiesService.getScriptProperties();
    var failCount = (parseInt(propsE.getProperty('CRAWL2_CONSEC_FAILURES') || '0') || 0) + 1;
    propsE.setProperty('CRAWL2_CONSEC_FAILURES', String(failCount));
    try {
      var es  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(propsE.getProperty('CRAWL2_ACTIVE_SHEET'));
      var elr = parseInt(propsE.getProperty('CRAWL2_LOG_ROW') || '0') || 0;
      if (failCount >= CRAWL2_MAX_CONSEC_FAILURES) {
        deleteCrawlV2Trigger();
        var msg = 'Error: ' + e.message.slice(0, 60);
        if (es)  setCrawlStatus(es, msg);
        if (elr) updateLogRow(elr, msg);
      } else {
        var retryMsg = 'Transient error (retry ' + failCount + '/' + CRAWL2_MAX_CONSEC_FAILURES + '): ' + e.message.slice(0, 60);
        if (es) setCrawlStatus(es, retryMsg);
      }
    } catch (e2) { /* swallow secondary error */ }
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// Public entry points (called from crawl_v2_panel.html via google.script.run)
// ============================================================

// seeds: optional array of hand-picked S2 paper objects (already resolved
// via findSeedPaper, reused from crawl.js) — added on top of the keyword
// pass's own discoveries, per the validated decision that the keyword pass
// replaces hand-picked seeds as the DEFAULT source but manual seeds remain
// addable as an explicit supplement.
function startCrawlV2(seeds, maxDepth, maxPapers, groups, crawlName, options) {
  try {
    var opts        = options || {};
    // Per-phase, not one global switch — venue sweep in particular is
    // exhaustive-by-design (§3) and can produce far more non-match noise
    // than keyword/backward/forward, so a venue-heavy run may want non-
    // matches recorded for keyword/backward/forward but not venue. Each
    // defaults to true (matches-only) if the panel doesn't send it.
    var matchesOnlyVenue    = opts.matchesOnlyVenue    !== false;
    var matchesOnlyKeyword  = opts.matchesOnlyKeyword  !== false;
    var matchesOnlyBackward = opts.matchesOnlyBackward !== false;
    var matchesOnlyForward  = opts.matchesOnlyForward  !== false;
    var yearBound   = opts.yearBound   !== false;
    // Backward's own depth, decoupled from forward's maxDepth — defaults to
    // 1 (only the original venue/keyword-found papers' own references are
    // examined; their references' references are not) so backward chasing
    // can't compound into ancient, generic citation noise. Forward keeps
    // exploring from whatever backward finds (including older foundational
    // work), up to the full maxDepth — that's how an older paper backward
    // turns up still generates forward links to modern citing work.
    var backwardMaxDepth = parseInt(opts.backwardMaxDepth) || 1;

    // Per-phase on/off toggles — reintroduced (v22 had removed them,
    // reasoning "the pipeline is the standard, not opt-in") specifically so
    // a keyword-only, year-bounded run can be tested in isolation before
    // deciding whether citation-chasing adds anything. Default true (today's
    // pipeline, unchanged) if the panel doesn't send a value. Unchecking
    // backward or forward also implies skipping the second backward pass —
    // see crawlV2BatchTrigger's phase-transition logic — since that pass
    // exists specifically to cover what forward finds.
    var runVenuePhase    = opts.runVenuePhase    !== false;
    var runBackwardPhase = opts.runBackwardPhase !== false;
    var runForwardPhase  = opts.runForwardPhase  !== false;
    var phase0Venues   = Array.isArray(opts.phase0Venues) ? opts.phase0Venues : [];
    var phase0YearFrom = parseInt(opts.phase0YearFrom) || PHASE0_YEAR_FROM_DEFAULT;
    var phase0YearTo   = parseInt(opts.phase0YearTo)   || PHASE0_YEAR_TO_DEFAULT;
    // Global term-match suppression list (v22 §6 — e.g. "from scratch"
    // shouldn't count as a "Scratch" hit) — not per filter group.
    var guardPhrases = Array.isArray(opts.guardPhrases) ? opts.guardPhrases : [];

    seeds = seeds || [];

    var seedYears = seeds
      .map(function(s) { return parseInt(s.year); })
      .filter(function(y) { return y && !isNaN(y); });
    var autoYearFloor    = seedYears.length ? Math.min.apply(null, seedYears) : 0;
    var yearFromOverride = parseInt(opts.yearFrom) || 0;
    var yearToOverride   = parseInt(opts.yearTo)   || 0;
    var yearFloor   = yearFromOverride || autoYearFloor;
    var yearCeiling = yearToOverride;

    var sheetName = (crawlName || '').trim() || ('v2 ' + newCrawlSheetName()); // newCrawlSheetName from crawl.js
    var seedLabel = seeds.length
      ? (seeds.length === 1 ? (seeds[0].title || 'Unknown') : seeds.length + ' hand-picked seed(s) + keyword pass')
      : 'Keyword pass (exhaustive boolean search)';

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.insertSheet(sheetName);
    setupCrawlV2Sheet(sheet, seedLabel);

    var startTime = new Date();
    sheet.getRange(1, CRAWL2_STARTED_VALUE_COL).setValue(formatTimestampV2(startTime));

    var props = PropertiesService.getScriptProperties();
    props.setProperty('CRAWL2_START_TIME',      startTime.toISOString());
    props.setProperty('CRAWL2_ACTIVE_SHEET',    sheetName);
    props.setProperty('CRAWL2_MAX_DEPTH',       String(maxDepth  || 2)); // depth 2 default (v22 §11.5 — was 3)
    props.setProperty('CRAWL2_BACKWARD_MAX_DEPTH', String(backwardMaxDepth));
    props.setProperty('CRAWL2_MAX_PAPERS',      String(maxPapers || 300));
    props.setProperty('CRAWL2_FILTER_GROUPS',   JSON.stringify(groups)); // v2's OWN key — not shared with v1/Snowball
    props.setProperty('CRAWL2_GUARD_PHRASES',   JSON.stringify(guardPhrases));
    // Phase sequence (venue -> keyword -> backward -> forward -> backward2
    // -> sweep) always starts at 'venue' — crawlV2BatchTrigger's own
    // per-phase handling skips straight past it (and past backward/
    // forward/backward2) if disabled below, so this doesn't need to change
    // based on which phases are enabled.
    props.setProperty('CRAWL2_PHASE',           'venue');
    props.setProperty('CRAWL2_RUN_VENUE_PHASE',    runVenuePhase    ? 'true' : 'false');
    props.setProperty('CRAWL2_RUN_BACKWARD_PHASE', runBackwardPhase ? 'true' : 'false');
    props.setProperty('CRAWL2_RUN_FORWARD_PHASE',  runForwardPhase  ? 'true' : 'false');
    props.setProperty('CRAWL2_MATCHES_ONLY_VENUE',    matchesOnlyVenue    ? 'true' : 'false');
    props.setProperty('CRAWL2_MATCHES_ONLY_KEYWORD',  matchesOnlyKeyword  ? 'true' : 'false');
    props.setProperty('CRAWL2_MATCHES_ONLY_BACKWARD', matchesOnlyBackward ? 'true' : 'false');
    props.setProperty('CRAWL2_MATCHES_ONLY_FORWARD',  matchesOnlyForward  ? 'true' : 'false');
    props.setProperty('CRAWL2_YEAR_BOUND',      yearBound   ? 'true' : 'false');
    props.setProperty('CRAWL2_YEAR_FLOOR',      String(yearFloor));
    props.setProperty('CRAWL2_YEAR_CEILING',    String(yearCeiling));
    props.setProperty('CRAWL2_KEYWORD_SUBQUERY_IDX',    '0');
    props.setProperty('CRAWL2_KEYWORD_TOKEN',           '');
    props.setProperty('CRAWL2_KEYWORD_PAGES_FETCHED',   '0');
    props.setProperty('CRAWL2_KEYWORD_COLLECTED',       '0');
    props.setProperty('CRAWL2_KEYWORD_RESULTS_SEEN',    '0');
    props.setProperty('CRAWL2_KEYWORD_SUBQUERY_ERRORS', '0');
    props.setProperty('CRAWL2_KEYWORD_TOTAL_ERRORS',    '0');
    props.setProperty('CRAWL2_KEYWORD_RETRIED_TERMS',   '[]');
    props.setProperty('CRAWL2_BACKWARD_IDX',      '0');
    props.setProperty('CRAWL2_BACKWARD_EXAMINED', '0');
    props.setProperty('CRAWL2_BACKWARD_KEPT',     '0');
    props.setProperty('CRAWL2_BACKWARD2_IDX',      '0');
    props.setProperty('CRAWL2_BACKWARD2_EXAMINED', '0');
    props.setProperty('CRAWL2_BACKWARD2_KEPT',     '0');
    props.setProperty('CRAWL2_SWEEP_IDX',         '0');
    props.setProperty('CRAWL2_SWEEP_RECOVERED',   '0');
    props.setProperty('CRAWL2_BATCH_NUM',         '1');
    props.setProperty('CRAWL2_SEEN_DOIS',         '[]'); // fresh cross-phase DOI dedup store for this crawl

    props.setProperty('CRAWL2_PHASE0_VENUES',    JSON.stringify(phase0Venues));
    props.setProperty('CRAWL2_PHASE0_YEAR_FROM', String(phase0YearFrom));
    props.setProperty('CRAWL2_PHASE0_YEAR_TO',   String(phase0YearTo));
    props.setProperty('CRAWL2_VENUE_BATCH_IDX',  '0');
    props.setProperty('CRAWL2_VENUE_TOKEN',      '');
    props.setProperty('CRAWL2_VENUE_COUNT',      '0');

    var logRow = appendLogRow('CrawlV2', {
      name:          sheetName,
      seeds:         seeds.map(function(s) {
        var mag = s.externalIds && s.externalIds.MAG;
        return mag ? ('W' + mag) : ('S2:' + s.paperId);
      }),
      depth:         maxDepth,
      maxPapers:     maxPapers,
      filterGroups:  groups,
      runBackward:   true,   // backward now always runs, both before and after forward
      expandBackward: true
    });
    if (logRow) props.setProperty('CRAWL2_LOG_ROW', String(logRow));

    // Write any hand-picked seeds immediately as depth-0 Direction='K' rows
    // — both hand-picked and keyword-found seeds share 'K' since both are
    // "phase 0/1 entry points" that predate any traversal, matching the
    // brief's 3-value (K/B/F) Direction acceptance criterion. The keyword
    // pass (once the trigger starts) adds more seeds on top of these.
    if (seeds.length > 0) {
      var seedRows = seeds.map(function(seed) { return crawlRowFromS2(seed, 0, '', 'K'); }); // from crawl.js
      sheet.getRange(CRAWL_FIRST_DATA_ROW, 1, seedRows.length, CRAWL_NUM_COLS).setValues(seedRows);
      sheet.getRange(CRAWL_FIRST_DATA_ROW, CRAWL_COL.CRAWLED, seedRows.length, 1).insertCheckboxes();
      sheet.setRowHeights(CRAWL_FIRST_DATA_ROW, seedRows.length, CRAWL_ROW_HEIGHT);
    }

    ss.setActiveSheet(sheet);
    applyCrawlV2Highlight(sheet, groups, guardPhrases); // tri-state K (TRUE/FALSE/REVIEW) + green/orange row highlight

    createCrawlV2Trigger();
    setCrawlStatus(sheet, 'Running venue sweep batch 1…');

    return 'Crawl v2 started — running in the background. Watch the "Crawl Status" cell at the top of the sheet. You can close this panel.';

  } catch (e) {
    return 'Error: ' + e.message;
  }
}

function resumeCrawlV2() {
  try {
    var props     = PropertiesService.getScriptProperties();
    var sheetName = props.getProperty('CRAWL2_ACTIVE_SHEET');
    if (!sheetName) return 'No active v2 crawl found. Start a new crawl first.';
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return 'Crawl sheet "' + sheetName + '" not found — it may have been deleted.';

    var phase = props.getProperty('CRAWL2_PHASE') || 'venue';

    // Only forward/backward/backward2 have a real Crawled=FALSE queue to
    // check for emptiness — venue/keyword/sweep track progress via their
    // own idx properties and are resumable regardless of countUncrawled().
    if ((phase === 'forward' || phase === 'backward' || phase === 'backward2') && countUncrawled(sheet) === 0) {
      return 'Crawl v2 "' + sheetName + '" has nothing left to process in this phase — it may already be complete.';
    }

    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
    var batch = parseInt(props.getProperty('CRAWL2_BATCH_NUM') || '1');
    createCrawlV2Trigger();
    setCrawlStatus(sheet, 'Resuming ' + phase + ' pass — batch ' + batch + '…');

    return 'Crawl v2 resumed — running in the background. Watch the "Crawl Status" cell at the top of the sheet.';
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

// Called from the v2 panel's "Apply Highlight Rule" button.
function applyCrawlV2Filter(groups, guardPhrases) {
  try {
    var props     = PropertiesService.getScriptProperties();
    var sheetName = props.getProperty('CRAWL2_ACTIVE_SHEET');
    if (!sheetName) return 'No active v2 crawl sheet found.';
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return 'Crawl sheet "' + sheetName + '" not found.';
    guardPhrases = Array.isArray(guardPhrases) ? guardPhrases : [];
    props.setProperty('CRAWL2_FILTER_GROUPS', JSON.stringify(groups));
    props.setProperty('CRAWL2_GUARD_PHRASES', JSON.stringify(guardPhrases));
    applyCrawlV2Highlight(sheet, groups, guardPhrases); // tri-state K (TRUE/FALSE/REVIEW) + green/orange row highlight
    return 'Highlight rule updated on "' + sheetName + '".';
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

// v2-only counterpart to snowball.js's getLastSnowballFilter() — guardPhrases
// had no restore path at all, so a fresh panel load always showed the
// textarea empty regardless of what was actually stored, and "Apply
// Highlight Rule" then persisted that emptiness over a working value (the
// bug: bare "Scratch" was matching inside "from scratch" because the guard
// phrase the user configured never actually reached the runtime). Reads the
// same CRAWL2_GUARD_PHRASES property startCrawlV2/applyCrawlV2Filter already
// write, so the panel's displayed value can never drift from what's
// actually stored for the active/most-recent crawl.
function getLastCrawlV2GuardPhrases() {
  var stored = PropertiesService.getScriptProperties().getProperty('CRAWL2_GUARD_PHRASES');
  if (!stored) return [];
  try {
    var parsed = JSON.parse(stored);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) { return []; }
}
