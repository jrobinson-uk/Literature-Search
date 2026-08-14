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

const CRAWL2_TIME_LIMIT_MS         = CRAWL_TIME_LIMIT_MS; // share v1's budget/reasoning
const CRAWL2_MAX_CONSEC_FAILURES   = CRAWL_MAX_CONSEC_FAILURES;

// Longer retry effort than v1's OPENALEX_BACKOFF_MS — validated response to
// "persist with the search for longer before moving on" rather than adding
// a third abstract source.
const OPENALEX_BACKOFF_MS_V2 = [1000, 2000, 4000, 8000, 15000, 30000];

const KEYWORD_SEARCH_TARGET_DEFAULT = 200;
const KEYWORD_RESULTS_PER_QUERY     = 10;
// Hard cap on generated queries so a filter with many terms per group can't
// runaway into thousands of API calls — the cartesian product across
// positive groups grows fast with term count.
const KEYWORD_MAX_QUERIES = 400;

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
  setupCrawlSheet(sheet, 'forward', seedLabel); // identical column layout to v1
  sheet.getRange(1, 1).setValue('Pipeline crawl (v2: keyword → backward → forward)');
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

function getCrawlV2ExistingKeys(sheet) {
  var lastRow = getCrawlLastDataRow(sheet);
  var ids     = new Set();
  var titles  = new Set();
  if (lastRow < 3) return { ids: ids, titles: titles };
  var idVals    = sheet.getRange(3, CRAWL_COL.ID,    lastRow - 2, 1).getValues().flat();
  var titleVals = sheet.getRange(3, CRAWL_COL.TITLE, lastRow - 2, 1).getValues().flat();
  idVals.forEach(function(v) { if (v) ids.add(v); });
  titleVals.forEach(function(v) {
    var norm = normalizeTitleV2(v);
    if (norm) titles.add(norm);
  });
  return { ids: ids, titles: titles };
}

// ============================================================
// Score-demotion filter — mirrors jsMatchesFilter's term-matching, but
// treats NOT-groups as advisory rather than a hard veto. A paper that
// passes every POSITIVE group is kept even if a NOT-group also matched —
// flagged via a title marker (same non-invasive convention as the existing
// fetch-failure / malformed-reference markers) instead of silently dropped.
// ============================================================

function termsAnyMatchV2(text, group) {
  if (!group.terms || !group.terms.trim()) return false;
  var terms = group.terms.split(",")
    .map(function(t) { return t.trim().replace(/^["']|["']$/g, "").trim().toLowerCase(); })
    .filter(function(t) { return t.length > 0; });
  if (terms.length === 0) return false;
  return terms.some(function(t) {
    return new RegExp('\\b' + escapeRegExpTerm(t) + '\\b', 'i').test(text); // escapeRegExpTerm from crawl.js
  });
}

// Returns { isMatch, flagged } — isMatch is true iff every positive group
// matched (NOT-groups don't affect it); flagged is true when isMatch is
// true AND a NOT-group also hit, i.e. exactly the class of paper the
// brief's Group-4-veto evidence showed being lost.
function jsMatchesFilterV2(text, groups) {
  text = (text || "").toLowerCase();
  var positive = groups.filter(function(g) { return !g.not; });
  var negative = groups.filter(function(g) { return g.not; });

  var positiveOk = positive.every(function(g) {
    if (!g.terms || !g.terms.trim()) return true;
    return termsAnyMatchV2(text, g);
  });
  var notHit = negative.some(function(g) { return termsAnyMatchV2(text, g); });

  return { isMatch: positiveOk, flagged: positiveOk && notHit };
}

// ============================================================
// Row builder — same shape as crawlRowFromS2, plus the score-demotion
// title marker.
// ============================================================

function buildCrawlV2Row(candidate, depth, parentId, dir, flagged) {
  var row = crawlRowFromS2(candidate, depth, parentId, dir); // from crawl.js
  if (flagged) {
    row[CRAWL_COL.TITLE - 1] = '⚠ [NOT-group override — review] ' + row[CRAWL_COL.TITLE - 1];
  }
  return row;
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
function buildKeywordQueries(groups) {
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

  var seen    = new Set();
  var queries = [];
  for (var i = 0; i < combos.length && queries.length < KEYWORD_MAX_QUERIES; i++) {
    var q = combos[i].join(' ');
    if (seen.has(q)) continue;
    seen.add(q);
    queries.push(q);
  }
  return queries;
}

// Resumable across trigger firings via CRAWL2_KEYWORD_IDX / _COLLECTED.
// Every kept seed is written Direction='K', Depth=0, Crawled=false — ready
// for the backward/forward phases to expand from once this phase ends.
function runKeywordPass(sheet, groups, targetSeeds, maxPapers, yearFloor, yearCeiling, yearBound) {
  var props   = PropertiesService.getScriptProperties();
  var queries = buildKeywordQueries(groups);
  var idx       = parseInt(props.getProperty('CRAWL2_KEYWORD_IDX')       || '0');
  var collected = parseInt(props.getProperty('CRAWL2_KEYWORD_COLLECTED') || '0');
  var startTime = Date.now();

  if (queries.length === 0) {
    return { status: 'complete', message: 'Keyword pass complete — no positive filter groups to search on. Add at least one non-NOT filter group.' };
  }

  var existing = getCrawlV2ExistingKeys(sheet);

  while (idx < queries.length && collected < targetSeeds) {
    if (Date.now() - startTime > CRAWL2_TIME_LIMIT_MS) {
      props.setProperty('CRAWL2_KEYWORD_IDX', String(idx));
      props.setProperty('CRAWL2_KEYWORD_COLLECTED', String(collected));
      return { status: 'time-limit', message: 'Keyword pass — time limit reached. ' + collected +
             ' seed(s) collected so far (' + idx + '/' + queries.length + ' queries run).' };
    }

    var query   = queries[idx++];
    var results = [];
    try { results = s2SearchPapers(query, KEYWORD_RESULTS_PER_QUERY); } catch (e) { /* skip a failed query */ }
    Utilities.sleep(1100); // same S2 pacing used elsewhere in the project

    var newRows  = [];
    var newNotes = [];
    results.forEach(function(paper) {
      if (!paper || !paper.paperId || collected >= targetSeeds) return;
      var mag = paper.externalIds && paper.externalIds.MAG;
      var id  = mag ? ('W' + mag) : ('S2:' + paper.paperId);
      if (existing.ids.has(id)) return;
      var normTitle = normalizeTitleV2(paper.title);
      if (normTitle && existing.titles.has(normTitle)) return;

      var abstract = paper.abstract || '';
      var note = null;
      if (!abstract) {
        var lookup = fetchOpenAlexAbstractV2(paper.externalIds);
        if (lookup.abstract) { abstract = lookup.abstract; paper.abstract = lookup.abstract; }
        note = describeAbstractSource(lookup); // from crawl.js
      }

      var yearOk  = isYearInBounds(paper.year, yearFloor, yearCeiling, yearBound); // from crawl.js
      var verdict = jsMatchesFilterV2((paper.title || '') + ' ' + abstract, groups);
      // The keyword pass only ever keeps genuine seeds — there's no "queue"
      // to record a dead-end non-match against the way forward/backward do,
      // since a rejected search result was never a candidate row to begin with.
      if (!verdict.isMatch || !yearOk) return;

      existing.ids.add(id);
      if (normTitle) existing.titles.add(normTitle);
      newRows.push(buildCrawlV2Row(paper, 0, '', 'K', verdict.flagged));
      newNotes.push(note);
      collected++;
    });

    if (newRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd        = Math.max(0, maxPapers - currentCount);
      var writeStart     = writeCrawlRows(sheet, newRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, newNotes.slice(0, canAdd));
    }
  }

  props.setProperty('CRAWL2_KEYWORD_IDX', String(idx));
  props.setProperty('CRAWL2_KEYWORD_COLLECTED', String(collected));
  return { status: 'complete', message: 'Keyword pass complete. ' + collected +
         ' seed paper(s) collected from ' + idx + ' quer' + (idx === 1 ? 'y' : 'ies') + '.' };
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
// ============================================================

function runBackwardPassV2(sheet, groups, maxDepth, maxPapers, yearFloor, yearCeiling, yearBound) {
  var startTime = Date.now();
  var props     = PropertiesService.getScriptProperties();
  var idx       = parseInt(props.getProperty('CRAWL2_BACKWARD_IDX') || '0');

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
    var verdict  = jsMatchesFilterV2(title + ' ' + abstract, groups);
    if (!verdict.isMatch) return; // only Filter Match=TRUE papers expand backward
    paperIds.push(id);
    paperSheetRows.push(3 + i);
    paperDepths.push(depth);
  });

  var processed = 0;
  while (idx < paperIds.length) {
    if (Date.now() - startTime > CRAWL2_TIME_LIMIT_MS) {
      props.setProperty('CRAWL2_BACKWARD_IDX', String(idx));
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

    var newRows  = [];
    var newNotes = [];
    refs.forEach(function(ref) {
      if (!ref || !ref.paperId) return;
      var mag   = ref.externalIds && ref.externalIds.MAG;
      var refId = mag ? ('W' + mag) : ('S2:' + ref.paperId);
      if (existing.ids.has(refId)) return;
      var normTitle = normalizeTitleV2(ref.title);
      if (normTitle && existing.titles.has(normTitle)) return;

      var note = null;
      if (!ref.abstract) {
        var lookup = fetchOpenAlexAbstractV2(ref.externalIds);
        if (lookup.abstract) ref.abstract = lookup.abstract;
        note = describeAbstractSource(lookup);
      }

      var yearOk  = isYearInBounds(ref.year, yearFloor, yearCeiling, yearBound);
      var verdict = jsMatchesFilterV2((ref.title || '') + ' ' + (ref.abstract || ''), groups);
      if (!verdict.isMatch || !yearOk) return;

      existing.ids.add(refId);
      if (normTitle) existing.titles.add(normTitle);
      newRows.push(buildCrawlV2Row(ref, paperDepth + 1, paperId, 'B', verdict.flagged));
      newNotes.push(note);
    });

    if (newRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd        = Math.max(0, maxPapers - currentCount);
      var writeStart     = writeCrawlRows(sheet, newRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, newNotes.slice(0, canAdd));
    }
  }

  props.setProperty('CRAWL2_BACKWARD_IDX', String(idx));
  updateCrawlMatchedCites(sheet);
  return { status: 'complete', message: 'Backward pass complete. Processed ' + processed + ' matching paper(s).' };
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

function runForwardPassV2(sheet, groups, maxDepth, maxPapers, matchesOnly, yearFloor, yearCeiling, yearBound) {
  matchesOnly = matchesOnly !== false;
  var startTime = Date.now();
  var processed  = 0;

  while (true) {
    if (Date.now() - startTime > CRAWL2_TIME_LIMIT_MS) {
      updateCrawlMatchedCites(sheet);
      var remaining = countUncrawled(sheet); // from crawl.js
      return { status: 'time-limit', message: "Forward pass — time limit reached. Processed " + processed +
             " papers this session; " + remaining + " remain in queue. Click Resume Crawl v2 to continue." };
    }

    var next = findNextUncrawled(sheet); // from crawl.js
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
      markFetchFailure(sheet, sheetRow, e);
      sheet.getRange(sheetRow, CRAWL_COL.CRAWLED).setValue(true);
      processed++;
      continue;
    }

    Utilities.sleep(1100);

    var existing      = getCrawlV2ExistingKeys(sheet);
    var matchRows     = [];
    var abstractNotes = [];
    for (var i = 0; i < candidates.length; i++) {
      var c   = candidates[i];
      var mag = c.externalIds && c.externalIds.MAG;
      var cId = mag ? ('W' + mag) : ('S2:' + c.paperId);
      if (existing.ids.has(cId)) continue;
      var normTitle = normalizeTitleV2(c.title);
      if (normTitle && existing.titles.has(normTitle)) continue;

      var abstract = c.abstract || '';
      var note = null;
      if (!abstract) {
        var lookup = fetchOpenAlexAbstractV2(c.externalIds);
        if (lookup.abstract) { abstract = lookup.abstract; c.abstract = lookup.abstract; }
        note = describeAbstractSource(lookup);
      }

      var yearOk  = isYearInBounds(getCandidateYear(c, "forward"), yearFloor, yearCeiling, yearBound); // from crawl.js
      var verdict = jsMatchesFilterV2((c.title || '') + ' ' + abstract, groups);
      var isMatch = verdict.isMatch && yearOk;
      if (!isMatch && matchesOnly) continue;
      var row = buildCrawlV2Row(c, depth + 1, id, 'F', verdict.flagged);
      if (!isMatch) row[CRAWL_COL.CRAWLED - 1] = true;
      matchRows.push(row);
      abstractNotes.push(note);
      existing.ids.add(cId);
      if (normTitle) existing.titles.add(normTitle);
    }

    if (matchRows.length > 0) {
      var currentCount = getCrawlLastDataRow(sheet) - 2;
      var canAdd        = Math.max(0, maxPapers - currentCount);
      var writeStart     = writeCrawlRows(sheet, matchRows.slice(0, canAdd));
      applyAbstractNotes(sheet, writeStart, abstractNotes.slice(0, canAdd));
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
// Phase order: keyword → backward (if enabled) → forward → sweep → complete.
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

    var phase       = props.getProperty('CRAWL2_PHASE') || 'keyword';
    var batch       = parseInt(props.getProperty('CRAWL2_BATCH_NUM') || '1');
    var groups      = JSON.parse(props.getProperty('CRAWL2_FILTER_GROUPS') || '[]');
    var maxDepth    = parseInt(props.getProperty('CRAWL2_MAX_DEPTH')  || '3');
    var maxPapers   = parseInt(props.getProperty('CRAWL2_MAX_PAPERS') || '300');
    var matchesOnly = props.getProperty('CRAWL2_MATCHES_ONLY') !== 'false';
    var runBackward = props.getProperty('CRAWL2_RUN_BACKWARD') !== 'false';
    var yearFloor   = parseInt(props.getProperty('CRAWL2_YEAR_FLOOR')   || '0') || 0;
    var yearCeiling = parseInt(props.getProperty('CRAWL2_YEAR_CEILING') || '0') || 0;
    var yearBound   = props.getProperty('CRAWL2_YEAR_BOUND') !== 'false';
    var targetSeeds = parseInt(props.getProperty('CRAWL2_TARGET_SEEDS') || String(KEYWORD_SEARCH_TARGET_DEFAULT));
    var logRow      = parseInt(props.getProperty('CRAWL2_LOG_ROW') || '0') || 0;

    var result;

    if (phase === 'keyword') {
      setCrawlStatus(sheet, 'Keyword pass — batch ' + batch + '…');
      result = runKeywordPass(sheet, groups, targetSeeds, maxPapers, yearFloor, yearCeiling, yearBound);
      if (result.status === 'time-limit') {
        props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, result.message);
      } else {
        props.setProperty('CRAWL2_PHASE', runBackward ? 'backward' : 'forward');
        props.setProperty('CRAWL2_BATCH_NUM', '1');
        setCrawlStatus(sheet, result.message + ' — starting ' + (runBackward ? 'backward' : 'forward') + ' pass…');
      }

    } else if (phase === 'backward') {
      setCrawlStatus(sheet, 'Backward pass — batch ' + batch + '…');
      result = runBackwardPassV2(sheet, groups, maxDepth, maxPapers, yearFloor, yearCeiling, yearBound);
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
      result = runForwardPassV2(sheet, groups, maxDepth, maxPapers, matchesOnly, yearFloor, yearCeiling, yearBound);
      if (result.status === 'time-limit') {
        props.setProperty('CRAWL2_BATCH_NUM', String(batch + 1));
        setCrawlStatus(sheet, result.message);
      } else if (result.status === 'paper-limit') {
        deleteCrawlV2Trigger();
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
    // Default ON — backward is core to the v2 pipeline (Phase 2), but still
    // toggleable off for a keyword+forward-only comparison run.
    var runBackward = opts.runBackward !== false;

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
    props.setProperty('CRAWL2_MAX_DEPTH',       String(maxDepth  || 3));
    props.setProperty('CRAWL2_MAX_PAPERS',      String(maxPapers || 300));
    props.setProperty('CRAWL2_TARGET_SEEDS',    String(targetN));
    props.setProperty('CRAWL2_FILTER_GROUPS',   JSON.stringify(groups)); // v2's OWN key — not shared with v1/Snowball
    props.setProperty('CRAWL2_PHASE',           'keyword');
    props.setProperty('CRAWL2_MATCHES_ONLY',    matchesOnly ? 'true' : 'false');
    props.setProperty('CRAWL2_RUN_BACKWARD',    runBackward ? 'true' : 'false');
    props.setProperty('CRAWL2_YEAR_BOUND',      yearBound   ? 'true' : 'false');
    props.setProperty('CRAWL2_YEAR_FLOOR',      String(yearFloor));
    props.setProperty('CRAWL2_YEAR_CEILING',    String(yearCeiling));
    props.setProperty('CRAWL2_KEYWORD_IDX',       '0');
    props.setProperty('CRAWL2_KEYWORD_COLLECTED', '0');
    props.setProperty('CRAWL2_BACKWARD_IDX',      '0');
    props.setProperty('CRAWL2_SWEEP_IDX',         '0');
    props.setProperty('CRAWL2_SWEEP_RECOVERED',   '0');
    props.setProperty('CRAWL2_BATCH_NUM',         '1');

    var logRow = appendLogRow('CrawlV2', {
      name:          sheetName,
      seeds:         seeds.map(function(s) {
        var mag = s.externalIds && s.externalIds.MAG;
        return mag ? ('W' + mag) : ('S2:' + s.paperId);
      }),
      depth:         maxDepth,
      maxPapers:     maxPapers,
      filterGroups:  groups,
      runBackward:   runBackward,
      expandBackward: false
    });
    if (logRow) props.setProperty('CRAWL2_LOG_ROW', String(logRow));

    // Write any hand-picked seeds immediately as depth-0 Direction='K' rows
    // — both hand-picked and keyword-found seeds share 'K' since both are
    // "phase 0/1 entry points" that predate any traversal, matching the
    // brief's 3-value (K/B/F) Direction acceptance criterion. The keyword
    // pass (once the trigger starts) adds more seeds on top of these.
    if (seeds.length > 0) {
      var seedRows = seeds.map(function(seed) { return buildCrawlV2Row(seed, 0, '', 'K', false); });
      sheet.getRange(3, 1, seedRows.length, CRAWL_NUM_COLS).setValues(seedRows);
      sheet.getRange(3, CRAWL_COL.CRAWLED, seedRows.length, 1).insertCheckboxes();
      sheet.setRowHeights(3, seedRows.length, CRAWL_ROW_HEIGHT);
    }

    ss.setActiveSheet(sheet);
    applyCrawlHighlight(sheet, groups); // reused directly from crawl.js — identical sheet layout

    createCrawlV2Trigger();
    setCrawlStatus(sheet, 'Running keyword pass batch 1…');

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

    var phase = props.getProperty('CRAWL2_PHASE') || 'keyword';
    // Only forward/backward have a real Crawled=FALSE queue to check for
    // emptiness — keyword and sweep track progress via their own idx
    // properties and are resumable regardless of countUncrawled().
    if ((phase === 'forward' || phase === 'backward') && countUncrawled(sheet) === 0) {
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
function applyCrawlV2Filter(groups) {
  try {
    var props     = PropertiesService.getScriptProperties();
    var sheetName = props.getProperty('CRAWL2_ACTIVE_SHEET');
    if (!sheetName) return 'No active v2 crawl sheet found.';
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return 'Crawl sheet "' + sheetName + '" not found.';
    props.setProperty('CRAWL2_FILTER_GROUPS', JSON.stringify(groups));
    applyCrawlHighlight(sheet, groups); // shared, identical sheet layout
    return 'Highlight rule updated on "' + sheetName + '".';
  } catch (e) {
    return 'Error: ' + e.message;
  }
}
