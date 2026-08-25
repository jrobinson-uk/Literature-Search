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

const KEYWORD_SEARCH_TARGET_DEFAULT = 200;
const KEYWORD_RESULTS_PER_QUERY     = 10;
// Default hard cap on generated queries for the plain (non-paginated)
// keyword pass — the cartesian product across positive groups grows fast
// with term count. The paginated alt-config raises this (see
// CRAWL2_PHASE1_MAX_QUERIES) since each query now pulls several pages
// instead of one, so more of the combination space can be reached before
// hitting a per-run API-call ceiling.
const KEYWORD_MAX_QUERIES         = 400;
const KEYWORD_MAX_QUERIES_PAGINATED = 800;

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
const PHASE1_PAGES_PER_QUERY_DEFAULT     = 3;
const PHASE1_SHORTFALL_TOLERANCE_DEFAULT = 0.5;

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
  if (lastRow < 3) return { ids: ids, titles: titles, dois: dois };
  var idVals    = sheet.getRange(3, CRAWL_COL.ID,    lastRow - 2, 1).getValues().flat();
  var titleVals = sheet.getRange(3, CRAWL_COL.TITLE, lastRow - 2, 1).getValues().flat();
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
function rememberCandidateV2(existing, id, normTitle, doi) {
  existing.ids.add(id);
  if (normTitle) existing.titles.add(normTitle);
  if (doi) {
    existing.dois.add(doi);
    saveSeenDois(existing.dois);
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
// Filter Match column (tri-state text) + row highlight
//
// Column K now holds "TRUE" / "FALSE" / "REVIEW" as literal text (v22 §2),
// computed by a sheet formula that mirrors jsMatchesFilterV2's rules
// exactly: positive groups gate first, then exclude-mode NOT groups, then
// review-mode NOT groups. The per-term helper columns feeding it apply the
// same guardPhrases masking as the JS side (buildTermFormulaV2), so the
// live display can't drift from the write-time decision the way two
// separately-evolving implementations could.
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
function buildTermFormulaV2(term, titleColNum, abstractColNum, guardPhrases) {
  const titleLetter  = colToLetter(titleColNum); // from snowball.js
  const absLetter    = colToLetter(abstractColNum);
  const safeTerm     = term.replace(/"/g, '""');
  const innerPattern = buildPluralAwareTermPattern(term, escapeTermForSheetFormula); // from crawl.js
  const maskedExpr   = buildMaskedTextExprV2('LOWER(t&" "&k)', guardPhrases);
  return '=MAP(A2:A,' + titleLetter + '2:' + titleLetter + ',' + absLetter + '2:' + absLetter + ',LAMBDA(a,t,k,' +
         'IF(ROW(a)=2,"' + safeTerm + '",' +
         'IF(a="","",' +
         'REGEXMATCH(' + maskedExpr + ',"\\b' + innerPattern + '\\b")' +
         '))))';
}

// Tri-state Filter Match formula: positive groups gate first (FALSE if any
// fails), then exclude-mode NOT groups (FALSE if any trips), then
// review-mode NOT groups (REVIEW if any trips), else TRUE — the exact same
// rule order as jsMatchesFilterV2.
function buildFilterMatchFormulaV2(parsedGroups, firstDetailColNum) {
  const totalTerms = parsedGroups.reduce(function(s, g) { return s + g.terms.length; }, 0);
  if (totalTerms === 0) return null;

  const params = parsedGroups.map(function(g, gi) {
    return g.terms.map(function(t, ti) {
      const idx = parsedGroups.slice(0, gi).reduce(function(s, pg) { return s + pg.terms.length; }, 0) + ti;
      return makeParamName(idx); // from snowball.js
    });
  });

  const ranges = params.map(function(groupParams, gi) {
    return groupParams.map(function(p, ti) {
      const colNum = firstDetailColNum +
        parsedGroups.slice(0, gi).reduce(function(s, pg) { return s + pg.terms.length; }, 0) + ti;
      const l = colToLetter(colNum); // from snowball.js
      return l + '2:' + l;
    }).join(',');
  }).join(',');

  const allParams = 'a,' + params.map(function(gp) { return gp.join(','); }).join(',');

  function orExprFor(gi) {
    const gParams = params[gi];
    return gParams.length === 1 ? gParams[0] : ('OR(' + gParams.join(',') + ')');
  }

  const positiveIdx = parsedGroups.map(function(g, gi) { return g.not ? null : gi; })
    .filter(function(x) { return x !== null; });
  const positiveExpr = positiveIdx.length === 0
    ? 'TRUE'
    : (positiveIdx.length === 1 ? orExprFor(positiveIdx[0]) : ('AND(' + positiveIdx.map(orExprFor).join(',') + ')'));

  const excludeIdx = parsedGroups.map(function(g, gi) { return (g.not && (g.notMode || 'exclude') === 'exclude') ? gi : null; })
    .filter(function(x) { return x !== null; });
  const excludeExpr = excludeIdx.length === 0
    ? 'FALSE'
    : (excludeIdx.length === 1 ? orExprFor(excludeIdx[0]) : ('OR(' + excludeIdx.map(orExprFor).join(',') + ')'));

  const reviewIdx = parsedGroups.map(function(g, gi) { return (g.not && g.notMode === 'review') ? gi : null; })
    .filter(function(x) { return x !== null; });
  const reviewExpr = reviewIdx.length === 0
    ? 'FALSE'
    : (reviewIdx.length === 1 ? orExprFor(reviewIdx[0]) : ('OR(' + reviewIdx.map(orExprFor).join(',') + ')'));

  const stateExpr =
    'IF(NOT(' + positiveExpr + '),"FALSE",IF(' + excludeExpr + ',"FALSE",IF(' + reviewExpr + ',"REVIEW","TRUE")))';

  return '=MAP(A2:A,' + ranges + ',LAMBDA(' + allParams + ',' +
         'IF(ROW(a)=2,"Filter Match",' +
         'IF(a="","",' + stateExpr + '))))';
}

// Replaces the shared applyCrawlHighlight (v1, archived) — identical term-
// helper column writing except it uses buildTermFormulaV2 (guardPhrases-
// aware) instead of the shared buildTermFormula, wires the tri-state
// formula above into Filter Match, and highlights rows by K's own text
// value directly (green=TRUE, orange=REVIEW) rather than needing a
// separate NOT-hit formula the way the pre-tri-state design did.
function applyCrawlV2Highlight(sheet, groups, guardPhrases) {
  var CRAWL_FIRST_DETAIL = CRAWL_FIRST_DETAIL_COL;

  sheet.getRange(2, CRAWL_IN_SHEET_LINKS_COL)
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
  sheet.getRange(2, CRAWL_FIRST_DETAIL, 1, maxHelper).clearContent().clearFormat();

  var colNum = CRAWL_FIRST_DETAIL;
  parsed.forEach(function(g, groupIdx) {
    var groupStartCol = colNum;
    var isNot  = g.not;
    var isReview = isNot && g.notMode === 'review';
    var termBg = isReview ? '#fff3cd' : (isNot ? '#fce8e6' : '#e8f0fe');

    g.terms.forEach(function(term) {
      sheet.getRange(2, colNum)
        .setFormula(buildTermFormulaV2(term, CRAWL_COL.TITLE, CRAWL_COL.ABSTRACT, guardPhrases))
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('bottom')
        .setTextRotation(90)
        .setBackground(termBg)
        .setFontColor('#222');
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
    sheet.setRowHeight(2, 130);

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

    // Row highlight now reads K's own tri-state text directly — no
    // separate NOT-hit formula needed, since K itself already distinguishes
    // TRUE/FALSE/REVIEW.
    var kLetter      = CRAWL_FILTER_MATCH_COL_LETTER; // "K" — from crawl.js
    var fullRowRange = sheet.getRange(3, 1, sheet.getMaxRows() - 2, CRAWL_NUM_COLS);
    var trueRowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + kLetter + '3="TRUE"')
      .setBackground('#b7e1cd') // green
      .setRanges([fullRowRange]).build();
    var reviewRowRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + kLetter + '3="REVIEW"')
      .setBackground('#ffe0b2') // orange
      .setRanges([fullRowRange]).build();

    // Drop anything touching the term-helper columns (stale per-term rules
    // from a previous Apply Highlight Rule click) or any prior row-
    // highlight rule keyed on K (both the old boolean =$K3=TRUE form and
    // the old score-demotion =AND($K3=TRUE,...) form, if this sheet was
    // created before the tri-state change) — keep everything else
    // (notably the no-abstract yellow-tint rule).
    var existingRules = sheet.getConditionalFormatRules().filter(function(rule) {
      if (rule.getRanges().some(function(r) { return r.getColumn() >= CRAWL_FIRST_DETAIL; })) return false;
      var bc = rule.getBooleanCondition();
      if (bc) {
        var vals    = bc.getCriteriaValues();
        var formula = vals && vals[0];
        if (formula && formula.indexOf('$' + kLetter + '3') !== -1) return false;
      }
      return true;
    });
    sheet.setConditionalFormatRules(existingRules.concat([trueRule, falseRule, trueRowRule, reviewRowRule]));
  }

  var filterFormula = buildFilterMatchFormulaV2(parsed, CRAWL_FIRST_DETAIL);
  if (!filterFormula) return;
  sheet.getRange(2, CRAWL_COL.FILTER_MATCH)
    .setFormula(filterFormula)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');
}

// Attaches a Flag Reason note to the Filter Match cell for a REVIEW row —
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
    sheet.getRange(startRow + i, CRAWL_COL.FILTER_MATCH).setNote(
      'Flagged for review: matches every positive filter group, but also ' +
      'tripped ' + groupLabel + termLabel + '. Kept and harvested, but not ' +
      'expanded — REVIEW rows are terminal nodes pending human triage.'
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

function runVenueSweep(sheet, groups, guardPhrases, venues, yearFrom, yearTo, maxPapers) {
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
      if (!verdict.isMatch) return; // FALSE — skip. REVIEW and TRUE both kept.

      rememberCandidateV2(existing, id, normTitle, doi);
      var row = crawlRowFromS2(paper, 0, '', 'V'); // from crawl.js — Direction='V'
      if (!verdict.expand) row[CRAWL_COL.CRAWLED - 1] = true; // REVIEW: harvested, never expanded
      newRows.push(row);
      newNotes.push(note);
      newFlags.push(verdict.state === 'REVIEW' ? { flagGroupIndex: verdict.flagGroupIndex, flagTerm: verdict.flagTerm } : null);
      collected++;
    });

    if (newRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd        = Math.max(0, maxPapers - currentCount);
      var writeStart     = writeCrawlRows(sheet, newRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, newNotes.slice(0, canAdd));
      applyFlagReasonNotes(sheet, writeStart, newFlags.slice(0, canAdd));
      if (canAdd < newRows.length) {
        props.setProperty('CRAWL2_VENUE_BATCH_IDX', String(batchIdx));
        props.setProperty('CRAWL2_VENUE_TOKEN', token);
        props.setProperty('CRAWL2_VENUE_COUNT', String(collected));
        return { status: 'paper-limit', message: 'Paper limit (' + maxPapers + ') reached during venue sweep — ' +
               'click Resume v2 Crawl to continue once you\'ve reviewed the sheet.' };
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

function s2SearchPapers(query, limit) {
  var fields = 'paperId,externalIds,title,abstract,year,authors,citationCount,publicationTypes,venue';
  var url    = 'https://api.semanticscholar.org/graph/v1/paper/search?query=' +
               encodeURIComponent(query) + '&fields=' + fields + '&limit=' + (limit || KEYWORD_RESULTS_PER_QUERY);
  var resp = s2Fetch(url); // from crawl.js — handles 429 back-off
  if (resp.getResponseCode() !== 200) return []; // one bad query shouldn't abort the whole pass
  var data = JSON.parse(resp.getContentText());
  return data.data || [];
}

// ============================================================
// S2 bulk-search — shared by Phase 0 (venue enumeration) and Phase 1's
// paginated alt-config. Distinct from s2SearchPapers's relevance-ranked
// /paper/search (which only ever returns a small top-N): /paper/search/bulk
// returns EVERYTHING matching the given filters, in pages of up to
// S2_BULK_PAGE_SIZE, via a continuation token — confirmed directly against
// the live endpoint before writing this (not assumed from memory):
//   - `query` is optional — omitting it entirely enumerates by venue/year
//     alone, which is exactly what Phase 0 needs ("the full publication
//     record", not a topical search).
//   - `venue` accepts a comma-separated list and fuzzy-matches each against
//     canonical venue names (e.g. "SIGCSE" alone matched "Technical
//     Symposium on Computer Science Education").
//   - `year` takes a "YYYY-YYYY" range.
//   - `sort` accepts "publicationDate:desc" for the date-sorted sweep.
//   - response is {total, token, data} — token is null once exhausted.
//
// opts: { query, venue, year, sort, token, fields }. Returns
// { total, token, data } — data is [] and token is null on any HTTP error
// (treated as "nothing more this call", same non-fatal-per-call convention
// as s2SearchPapers) rather than throwing, so one bad page doesn't abort
// an entire venue/query sweep.
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
  var resp = s2Fetch(url); // from crawl.js — handles 429 back-off
  if (resp.getResponseCode() !== 200) return { total: 0, token: null, data: [] };
  var data = JSON.parse(resp.getContentText());
  if (data.error) return { total: 0, token: null, data: [] }; // e.g. an over-broad query with no venue/year filter
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

// Builds the query strings the keyword pass searches — one term chosen per
// positive filter group, so each query is a natural-language phrase (S2
// search has no boolean operators; it's relevance-ranked, not literal AND).
// NOT-groups don't contribute to query generation, only to post-filtering,
// same as everywhere else. Capped at KEYWORD_MAX_QUERIES.
// maxQueries defaults to KEYWORD_MAX_QUERIES (plain mode) but the
// paginated alt-config passes a higher cap (KEYWORD_MAX_QUERIES_PAGINATED)
// since each query now pulls several pages instead of one.
//
// Combos are shuffled before the cap is applied — reduce()'s nested-loop
// order otherwise means whichever positive group is listed FIRST gets full
// term coverage while later groups' terms are systematically under-
// represented once the list is truncated (a concrete, code-level
// explanation for the brief's Gap 1 "generic terms absorbing the budget"
// complaint: if a broad term sits early in group order, every combination
// containing it survives a truncation that drops later, narrower combos).
function buildKeywordQueries(groups, maxQueries) {
  maxQueries = maxQueries || KEYWORD_MAX_QUERIES;
  var positiveTermLists = groups
    .filter(function(g) { return !g.not; })
    .map(parseGroupTermsV2)
    .filter(function(list) { return list.length > 0; });

  if (positiveTermLists.length === 0) return [];

  var combos = positiveTermLists.reduce(function(acc, termList) {
    var next = [];
    acc.forEach(function(prefix) {
      termList.forEach(function(term) { next.push(prefix.concat([term])); });
    });
    return next;
  }, [[]]);

  for (var i = combos.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = combos[i]; combos[i] = combos[j]; combos[j] = tmp;
  }

  var seen    = new Set();
  var queries = [];
  for (var k = 0; k < combos.length && queries.length < maxQueries; k++) {
    var q = combos[k].join(' ');
    if (seen.has(q)) continue;
    seen.add(q);
    queries.push(q);
  }
  return queries;
}

// Resumable across trigger firings via CRAWL2_KEYWORD_IDX / _COLLECTED.
// Every kept seed is written Direction='K', Depth=0, Crawled=false — ready
// for the backward/forward phases to expand from once this phase ends.
//
// opts (all off/default unless set, so a plain call reproduces today's
// v19 behaviour exactly — v3 §0.2):
//   paginated:          use S2's bulk-search endpoint with up to
//                        pagesPerQuery pages per query, instead of a
//                        single small page from the relevance-only
//                        /paper/search endpoint.
//   pagesPerQuery:       default PHASE1_PAGES_PER_QUERY_DEFAULT (3).
//   dateSweep:           after the relevance-sorted pass exhausts its
//                        query list, run the SAME queries again sorted by
//                        publicationDate:desc — only meaningful (and only
//                        applied) when paginated is also on, since only
//                        the bulk endpoint exposes a sort param.
//   noYearFloor:         ignore yearBound entirely for this phase's own
//                        candidates (backward/forward keep it regardless).
//   maxQueries:          query-generation cap; defaults to
//                        KEYWORD_MAX_QUERIES_PAGINATED when paginated,
//                        else KEYWORD_MAX_QUERIES.
//   shortfallTolerance:  fraction of targetSeeds below which the phase
//                        returns status 'shortfall' instead of 'complete'
//                        — required instrumentation per the brief ("a 14%
//                        shortfall passed silently in v19"); the trigger
//                        stops rather than cascading into backward/forward
//                        with a thin seed set.
function runKeywordPass(sheet, groups, guardPhrases, targetSeeds, maxPapers, yearFloor, yearCeiling, yearBound, opts) {
  opts = opts || {};
  // v22 §4/§0.1: pagination, the date-sorted sweep, and no year floor are
  // now the default behaviour of Phase 1 (not opt-in flags) — each still
  // defaults true even if the caller passes no opts at all. Only pass
  // `false` explicitly to turn one off (kept for flexibility/testing, not
  // exposed as a toggle in the panel any more).
  var paginated  = opts.paginated !== false;
  var pagesPerQuery = parseInt(opts.pagesPerQuery) || PHASE1_PAGES_PER_QUERY_DEFAULT;
  var dateSweep  = opts.dateSweep !== false && paginated;
  var noYearFloor = opts.noYearFloor !== false;
  var maxQueries = parseInt(opts.maxQueries) || (paginated ? KEYWORD_MAX_QUERIES_PAGINATED : KEYWORD_MAX_QUERIES);
  var shortfallTolerance = (opts.shortfallTolerance != null && opts.shortfallTolerance !== '')
    ? parseFloat(opts.shortfallTolerance) : PHASE1_SHORTFALL_TOLERANCE_DEFAULT;
  var effectiveYearBound = noYearFloor ? false : yearBound;

  var props   = PropertiesService.getScriptProperties();
  var queries = buildKeywordQueries(groups, maxQueries);
  var subphase      = props.getProperty('CRAWL2_KEYWORD_SUBPHASE')        || 'relevance'; // 'relevance' | 'date'
  var idx           = parseInt(props.getProperty('CRAWL2_KEYWORD_IDX')             || '0');
  var collected     = parseInt(props.getProperty('CRAWL2_KEYWORD_COLLECTED')       || '0');
  var queriesIssued = parseInt(props.getProperty('CRAWL2_KEYWORD_QUERIES_ISSUED')  || '0');
  var resultsSeen   = parseInt(props.getProperty('CRAWL2_KEYWORD_RESULTS_SEEN')    || '0');
  var startTime = Date.now();

  if (queries.length === 0) {
    return { status: 'complete', message: 'Keyword pass complete — no positive filter groups to search on. Add at least one non-NOT filter group.' };
  }

  var existing = getCrawlV2ExistingKeys(sheet);

  function persistProgress() {
    props.setProperty('CRAWL2_KEYWORD_SUBPHASE', subphase);
    props.setProperty('CRAWL2_KEYWORD_IDX', String(idx));
    props.setProperty('CRAWL2_KEYWORD_COLLECTED', String(collected));
    props.setProperty('CRAWL2_KEYWORD_QUERIES_ISSUED', String(queriesIssued));
    props.setProperty('CRAWL2_KEYWORD_RESULTS_SEEN', String(resultsSeen));
  }

  while (collected < targetSeeds) {
    if (idx >= queries.length) {
      // Relevance sweep exhausted — hand off to the date sweep (same query
      // list, re-sorted) if enabled, otherwise this phase is done.
      if (dateSweep && subphase === 'relevance') {
        subphase = 'date';
        idx = 0;
        continue;
      }
      break;
    }

    if (Date.now() - startTime > CRAWL2_TIME_LIMIT_MS) {
      persistProgress();
      return { status: 'time-limit', message: 'Keyword pass (' + subphase + ' sweep) — time limit reached. ' +
             collected + ' seed(s) collected so far (' + idx + '/' + queries.length + ' queries this sweep).' };
    }

    var query   = queries[idx++];
    var results = [];
    if (paginated) {
      var token = null;
      for (var page = 0; page < pagesPerQuery; page++) {
        var pageResult = s2BulkSearch({
          query: query,
          sort:  subphase === 'date' ? 'publicationDate:desc' : null,
          token: token
        });
        queriesIssued++;
        results = results.concat(pageResult.data);
        if (!pageResult.token) break;
        token = pageResult.token;
        Utilities.sleep(1100);
      }
    } else {
      try { results = s2SearchPapers(query, KEYWORD_RESULTS_PER_QUERY); } catch (e) { /* skip a failed query */ }
      queriesIssued++;
    }
    resultsSeen += results.length;
    Utilities.sleep(1100); // same S2 pacing used elsewhere in the project

    var newRows  = [];
    var newNotes = [];
    var newFlags = []; // aligned with newRows — flag info {flagGroupIndex, flagTerm} or null
    results.forEach(function(paper) {
      if (!paper || !paper.paperId || collected >= targetSeeds) return;
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
      // The keyword pass only ever keeps genuine seeds — there's no "queue"
      // to record a dead-end non-match against the way forward/backward do,
      // since a rejected search result was never a candidate row to begin with.
      if (!verdict.isMatch || !yearOk) return; // FALSE, or year-rejected — skip. REVIEW and TRUE both kept.

      rememberCandidateV2(existing, id, normTitle, doi);
      var row = crawlRowFromS2(paper, 0, '', 'K'); // from crawl.js
      if (!verdict.expand) row[CRAWL_COL.CRAWLED - 1] = true; // REVIEW: harvested, never expanded
      newRows.push(row);
      newNotes.push(note);
      newFlags.push(verdict.state === 'REVIEW' ? { flagGroupIndex: verdict.flagGroupIndex, flagTerm: verdict.flagTerm } : null);
      collected++;
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

  var summary = 'Keyword pass complete. ' + collected + ' seed paper(s) collected from ' +
    queriesIssued + ' quer' + (queriesIssued === 1 ? 'y call' : 'y calls') + ' (' + resultsSeen + ' results seen)' +
    (dateSweep ? ', relevance + date sweeps' : '') + '.';

  // Required instrumentation, not optional (the brief's own framing): fail
  // loudly rather than silently proceeding when the shortfall is large —
  // a 14% yield (280/2000) passed silently in v19 straight into backward/forward.
  if (collected < targetSeeds * shortfallTolerance) {
    return { status: 'shortfall', message: 'Keyword pass shortfall: only ' + collected + ' of ' + targetSeeds +
           ' target seeds collected (' + Math.round(100 * collected / targetSeeds) + '%, below the ' +
           Math.round(100 * shortfallTolerance) + '% tolerance) from ' + queriesIssued + ' quer' +
           (queriesIssued === 1 ? 'y call' : 'y calls') + ' (' + resultsSeen + ' results seen). Stopped rather than ' +
           'continuing to backward/forward with a thin seed set — review the filter groups or query construction, ' +
           'then click Resume v2 Crawl to continue anyway.' };
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
function runBackwardPassV2(sheet, groups, guardPhrases, maxDepth, maxPapers, yearFloor, yearCeiling, yearBound, propPrefix) {
  propPrefix = propPrefix || 'CRAWL2_BACKWARD';
  var startTime = Date.now();
  var props     = PropertiesService.getScriptProperties();
  var idx       = parseInt(props.getProperty(propPrefix + '_IDX') || '0');
  // Reject-counting (v3 §4: "a per-parent count of references examined vs
  // kept, so the phase has a measurable hit rate") — aggregate, not
  // per-parent-logged, per the brief's own "at minimum" allowance.
  var examined  = parseInt(props.getProperty(propPrefix + '_EXAMINED') || '0');
  var kept      = parseInt(props.getProperty(propPrefix + '_KEPT')     || '0');

  var lastRow = getCrawlLastDataRow(sheet);
  if (lastRow < 3) return { status: 'complete', message: 'Backward pass complete. No papers found.' };

  var numRows = lastRow - 2;
  var data    = sheet.getRange(3, 1, numRows, CRAWL_NUM_COLS).getValues();

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
    paperSheetRows.push(3 + i);
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
      if (!verdict.isMatch || !yearOk) return; // FALSE, or year-rejected — skip. REVIEW and TRUE both kept.

      rememberCandidateV2(existing, refId, normTitle, doi);
      var row = crawlRowFromS2(ref, paperDepth + 1, paperId, 'B'); // from crawl.js
      if (!verdict.expand) row[CRAWL_COL.CRAWLED - 1] = true; // REVIEW: harvested, never expanded
      newRows.push(row);
      newNotes.push(note);
      newFlags.push(verdict.state === 'REVIEW' ? { flagGroupIndex: verdict.flagGroupIndex, flagTerm: verdict.flagTerm } : null);
      kept++;
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
      rememberCandidateV2(existing, cId, normTitle, cDoi);
    }

    if (matchRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd        = Math.max(0, maxPapers - currentCount);
      var writeStart     = writeCrawlRows(sheet, matchRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, abstractNotes.slice(0, canAdd));
      applyFlagReasonNotes(sheet, writeStart, flagInfos.slice(0, canAdd));
      if (canAdd < matchRows.length) {
        sheet.getRange(sheetRow, CRAWL_COL.CRAWLED).setValue(true);
        processed++;
        updateCrawlMatchedCites(sheet);
        var logRowPL = parseInt(PropertiesService.getScriptProperties().getProperty('CRAWL2_LOG_ROW') || '0') || 0;
        updateLogRow(logRowPL, 'Paper Limit');
        return { status: 'paper-limit', message: "Paper limit (" + maxPapers + ") reached. The queue still has unprocessed papers — " +
               "click Resume Crawl v2 to continue (existing queue only; no new papers will be added)." };
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
  if (lastRow < 3) return { status: 'complete', recovered: 0, nextIdx: 0 };
  var numRows = lastRow - 2;
  var idCol   = sheet.getRange(3, CRAWL_COL.ID,       numRows, 1).getValues();
  var absCol  = sheet.getRange(3, CRAWL_COL.ABSTRACT, numRows, 1).getValues();
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
    var row = 3 + i;
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
    var batch       = parseInt(props.getProperty('CRAWL2_BATCH_NUM') || '1');
    var groups      = JSON.parse(props.getProperty('CRAWL2_FILTER_GROUPS') || '[]');
    var guardPhrases = JSON.parse(props.getProperty('CRAWL2_GUARD_PHRASES') || '[]');
    var maxDepth    = parseInt(props.getProperty('CRAWL2_MAX_DEPTH')  || '2');
    var maxPapers   = parseInt(props.getProperty('CRAWL2_MAX_PAPERS') || '300');
    var matchesOnly = props.getProperty('CRAWL2_MATCHES_ONLY') !== 'false';
    var yearFloor   = parseInt(props.getProperty('CRAWL2_YEAR_FLOOR')   || '0') || 0;
    var yearCeiling = parseInt(props.getProperty('CRAWL2_YEAR_CEILING') || '0') || 0;
    var yearBound   = props.getProperty('CRAWL2_YEAR_BOUND') !== 'false';
    var targetSeeds = parseInt(props.getProperty('CRAWL2_TARGET_SEEDS') || String(KEYWORD_SEARCH_TARGET_DEFAULT));
    var logRow      = parseInt(props.getProperty('CRAWL2_LOG_ROW') || '0') || 0;

    var phase0Venues   = JSON.parse(props.getProperty('CRAWL2_PHASE0_VENUES') || '[]');
    var phase0YearFrom = parseInt(props.getProperty('CRAWL2_PHASE0_YEAR_FROM') || String(PHASE0_YEAR_FROM_DEFAULT));
    var phase0YearTo   = parseInt(props.getProperty('CRAWL2_PHASE0_YEAR_TO')   || String(PHASE0_YEAR_TO_DEFAULT));
    var keywordOpts = {
      pagesPerQuery:      parseInt(props.getProperty('CRAWL2_PHASE1_PAGES_PER_QUERY') || String(PHASE1_PAGES_PER_QUERY_DEFAULT)),
      maxQueries:         parseInt(props.getProperty('CRAWL2_PHASE1_MAX_QUERIES') || '0') || 0,
      shortfallTolerance: parseFloat(props.getProperty('CRAWL2_PHASE1_SHORTFALL_TOLERANCE') || String(PHASE1_SHORTFALL_TOLERANCE_DEFAULT))
      // paginated/dateSweep/noYearFloor deliberately omitted — runKeywordPass
      // defaults all three to true now that they're the standard behaviour,
      // not opt-in flags.
    };

    var result;

    if (phase === 'venue') {
      setCrawlStatus(sheet, 'Venue sweep — batch ' + batch + '…');
      result = runVenueSweep(sheet, groups, guardPhrases, phase0Venues, phase0YearFrom, phase0YearTo, maxPapers);
      if (result.status === 'time-limit') {
        props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, result.message);
      } else if (result.status === 'paper-limit') {
        deleteCrawlV2Trigger();
        setCrawlStatus(sheet, result.message);
      } else {
        props.setProperty('CRAWL2_PHASE', 'keyword');
        props.setProperty('CRAWL2_BATCH_NUM', '1');
        setCrawlStatus(sheet, result.message + ' — starting keyword pass…');
      }

    } else if (phase === 'keyword') {
      setCrawlStatus(sheet, 'Keyword pass — batch ' + batch + '…');
      result = runKeywordPass(sheet, groups, guardPhrases, targetSeeds, maxPapers, yearFloor, yearCeiling, yearBound, keywordOpts);
      if (result.status === 'time-limit') {
        props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, result.message);
      } else if (result.status === 'shortfall') {
        // Required instrumentation (§4): stop rather than cascade into
        // backward/forward with a thin seed set — resumable, same
        // "Resume v2 Crawl continues anyway" convention as Paper Limit.
        // CRAWL2_KEYWORD_SHORTFALL marks *why* it stopped, so resumeCrawlV2
        // can tell "acknowledge and move to backward" apart from "genuinely
        // still has more queries to run" — without it, Resume just re-enters
        // this same phase, re-evaluates the identical (already exhausted)
        // query state, and immediately reports the same shortfall again.
        props.setProperty('CRAWL2_KEYWORD_SHORTFALL', 'true');
        deleteCrawlV2Trigger();
        updateLogRow(logRow, 'Shortfall');
        setCrawlStatus(sheet, result.message);
      } else {
        props.setProperty('CRAWL2_PHASE', 'backward');
        props.setProperty('CRAWL2_BATCH_NUM', '1');
        setCrawlStatus(sheet, result.message + ' — starting backward pass…');
      }

    } else if (phase === 'backward') {
      setCrawlStatus(sheet, 'Backward pass — batch ' + batch + '…');
      result = runBackwardPassV2(sheet, groups, guardPhrases, maxDepth, maxPapers, yearFloor, yearCeiling, yearBound, 'CRAWL2_BACKWARD');
      if (result.status === 'time-limit') {
        props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, result.message);
      } else {
        props.setProperty('CRAWL2_PHASE', 'forward');
        props.setProperty('CRAWL2_BATCH_NUM', '1');
        setCrawlStatus(sheet, result.message + ' — starting forward pass…');
      }

    } else if (phase === 'forward') {
      setCrawlStatus(sheet, 'Forward pass — batch ' + batch + '…');
      result = runForwardPassV2(sheet, groups, guardPhrases, maxDepth, maxPapers, matchesOnly, yearFloor, yearCeiling, yearBound);
      if (result.status === 'time-limit') {
        props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, result.message);
      } else if (result.status === 'paper-limit') {
        deleteCrawlV2Trigger();
        setCrawlStatus(sheet, result.message);
      } else {
        props.setProperty('CRAWL2_PHASE', 'backward2');
        props.setProperty('CRAWL2_BATCH_NUM', '1');
        setCrawlStatus(sheet, result.message + ' — starting second backward pass (over the now-larger match set)…');
      }

    } else if (phase === 'backward2') {
      setCrawlStatus(sheet, 'Second backward pass — batch ' + batch + '…');
      result = runBackwardPassV2(sheet, groups, guardPhrases, maxDepth, maxPapers, yearFloor, yearCeiling, yearBound, 'CRAWL2_BACKWARD2');
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
function startCrawlV2(seeds, maxDepth, maxPapers, targetSeeds, groups, crawlName, options) {
  try {
    var opts        = options || {};
    var matchesOnly = opts.matchesOnly !== false;
    var yearBound   = opts.yearBound   !== false;

    // v22 §0.1/§1: venue sweep, the backward dual-pass, and Phase 1's
    // pagination/date-sweep/no-floor are now the standard pipeline, not
    // opt-in flags — no on/off toggles for any of these any more. Their
    // own tunable parameters (venue list, year window, pages/query, max
    // queries, shortfall tolerance) remain configurable.
    var phase0Venues   = Array.isArray(opts.phase0Venues) ? opts.phase0Venues : [];
    var phase0YearFrom = parseInt(opts.phase0YearFrom) || PHASE0_YEAR_FROM_DEFAULT;
    var phase0YearTo   = parseInt(opts.phase0YearTo)   || PHASE0_YEAR_TO_DEFAULT;
    var phase1PagesPerQuery      = parseInt(opts.phase1PagesPerQuery) || PHASE1_PAGES_PER_QUERY_DEFAULT;
    var phase1MaxQueries         = parseInt(opts.phase1MaxQueries) || 0; // 0 = let runKeywordPass pick its own default
    var phase1ShortfallTolerance = (opts.phase1ShortfallTolerance != null && opts.phase1ShortfallTolerance !== '')
      ? parseFloat(opts.phase1ShortfallTolerance) : PHASE1_SHORTFALL_TOLERANCE_DEFAULT;
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
    var targetN   = parseInt(targetSeeds) || KEYWORD_SEARCH_TARGET_DEFAULT;
    var seedLabel = seeds.length
      ? (seeds.length === 1 ? (seeds[0].title || 'Unknown') : seeds.length + ' hand-picked seed(s) + keyword pass')
      : 'Keyword pass (target ' + targetN + ' seeds)';

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.insertSheet(sheetName);
    setupCrawlV2Sheet(sheet, seedLabel);

    var props = PropertiesService.getScriptProperties();
    props.setProperty('CRAWL2_ACTIVE_SHEET',    sheetName);
    props.setProperty('CRAWL2_MAX_DEPTH',       String(maxDepth  || 2)); // depth 2 default (v22 §11.5 — was 3)
    props.setProperty('CRAWL2_MAX_PAPERS',      String(maxPapers || 300));
    props.setProperty('CRAWL2_TARGET_SEEDS',    String(targetN));
    props.setProperty('CRAWL2_FILTER_GROUPS',   JSON.stringify(groups)); // v2's OWN key — not shared with v1/Snowball
    props.setProperty('CRAWL2_GUARD_PHRASES',   JSON.stringify(guardPhrases));
    // Fixed phase sequence now (venue -> keyword -> backward -> forward ->
    // backward2 -> sweep) — venue always starts first; it's a harmless
    // no-op if no venues are configured.
    props.setProperty('CRAWL2_PHASE',           'venue');
    props.setProperty('CRAWL2_MATCHES_ONLY',    matchesOnly ? 'true' : 'false');
    props.setProperty('CRAWL2_YEAR_BOUND',      yearBound   ? 'true' : 'false');
    props.setProperty('CRAWL2_YEAR_FLOOR',      String(yearFloor));
    props.setProperty('CRAWL2_YEAR_CEILING',    String(yearCeiling));
    props.setProperty('CRAWL2_KEYWORD_IDX',       '0');
    props.setProperty('CRAWL2_KEYWORD_COLLECTED', '0');
    props.setProperty('CRAWL2_KEYWORD_SUBPHASE',        'relevance');
    props.setProperty('CRAWL2_KEYWORD_QUERIES_ISSUED',  '0');
    props.setProperty('CRAWL2_KEYWORD_RESULTS_SEEN',    '0');
    // Script-wide property, not scoped to a single crawl sheet — without
    // resetting it here, a fresh crawl could inherit 'true' left over from
    // an earlier crawl's shortfall and have Resume misfire straight past
    // its own (first, legitimate) keyword pass.
    props.deleteProperty('CRAWL2_KEYWORD_SHORTFALL');
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
    props.setProperty('CRAWL2_PHASE1_PAGES_PER_QUERY',      String(phase1PagesPerQuery));
    props.setProperty('CRAWL2_PHASE1_MAX_QUERIES',          String(phase1MaxQueries));
    props.setProperty('CRAWL2_PHASE1_SHORTFALL_TOLERANCE',  String(phase1ShortfallTolerance));

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
      sheet.getRange(3, 1, seedRows.length, CRAWL_NUM_COLS).setValues(seedRows);
      sheet.getRange(3, CRAWL_COL.CRAWLED, seedRows.length, 1).insertCheckboxes();
      sheet.setRowHeights(3, seedRows.length, CRAWL_ROW_HEIGHT);
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

    // A shortfall stop means the keyword pass already exhausted every query
    // it had (both sweeps) — re-entering 'keyword' as-is would just
    // re-evaluate that same exhausted state and immediately report the
    // identical shortfall again. Resuming here means "continue anyway", so
    // skip straight to backward with whatever was collected.
    if (phase === 'keyword' && props.getProperty('CRAWL2_KEYWORD_SHORTFALL') === 'true') {
      props.deleteProperty('CRAWL2_KEYWORD_SHORTFALL');
      props.setProperty('CRAWL2_PHASE', 'backward');
      props.setProperty('CRAWL2_BATCH_NUM', '1');
      phase = 'backward';
    }

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
