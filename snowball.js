// Column layout (1-based)
// Round | Include | Fetched | Ref Count | Cited By | Year | Title | Authors | Type | Venue | Abstract | ID | Cites
// A1 = current round counter | B1 = label
// Row 2 = headers
// Data from row 3

const HEADERS  = ["Round", "Include", "Fetched", "Ref Count", "Cited By", "Year", "Title", "Authors", "Type", "Venue", "Abstract", "ID", "Cites", "Filter Match"];
const NUM_COLS = HEADERS.length;

const COL = {
  ROUND:        1,
  INCLUDE:      2,
  FETCHED:      3,
  REF_COUNT:    4,
  CITED_BY:     5,
  YEAR:         6,
  TITLE:        7,
  AUTHORS:      8,
  PUB_TYPES:    9,
  VENUE:        10,
  ABSTRACT:     11,
  ID:           12,
  CITES:        13,
  FILTER_MATCH: 14
};

// Column letter for FILTER_MATCH (col 14 = N) — used in CF formula
const FILTER_MATCH_COL_LETTER = "N";
// Default MAP formula when no filter has been applied
const DEFAULT_FILTER_FORMULA =
  '=MAP(A2:A,K2:K,LAMBDA(a,k,IF(ROW(a)=2,"Filter Match",IF(a<>"",FALSE,""))))';

const ROW_HEIGHT = 75;

// First column used for per-group detail formulas (immediately right of Filter Match)
const FIRST_DETAIL_COL = COL.FILTER_MATCH + 1; // col 15 = O

// ---------------------------------------------------------------------------
// Formula helpers
// ---------------------------------------------------------------------------

// Converts 1-based column number to A1 letter notation (A, B, ..., Z, AA, ...)
function colToLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// Generates a safe LAMBDA parameter name for index i (0-based).
// Uses letters b-z (25), then bb-bz, cb-cz, … — never digit-suffixed,
// so Sheets never confuses them with cell references like B1, G3, etc.
function makeParamName(i) {
  const alpha = 'bcdefghijklmnopqrstuvwxyz'; // 25 letters, no 'a'
  if (i < 25) return alpha[i];
  const j = i - 25;
  return alpha[Math.floor(j / 25)] + alpha[j % 25]; // bb, bc … bz, cb, cc …
}

// Builds the Filter Match MAP formula by referencing the per-term columns.
// For each non-NOT group: OR across that group's columns must be TRUE.
// For each NOT group:    OR across that group's columns must be FALSE.
function buildFilterMatchFromTermCols(parsedGroups, firstDetailColNum) {
  const totalTerms = parsedGroups.reduce(function(s, g) { return s + g.terms.length; }, 0);
  if (totalTerms === 0) return null;

  const params = parsedGroups.map(function(g, gi) {
    return g.terms.map(function(t, ti) {
      const idx = parsedGroups.slice(0, gi).reduce(function(s, pg) { return s + pg.terms.length; }, 0) + ti;
      return makeParamName(idx);
    });
  });

  const ranges = params.map(function(groupParams, gi) {
    return groupParams.map(function(p, ti) {
      const colNum = firstDetailColNum +
        parsedGroups.slice(0, gi).reduce(function(s, pg) { return s + pg.terms.length; }, 0) + ti;
      const l = colToLetter(colNum);
      return l + '2:' + l;
    }).join(',');
  }).join(',');

  const allParams = 'a,' + params.map(function(gp) { return gp.join(','); }).join(',');

  const andParts = parsedGroups.map(function(g, gi) {
    const gParams = params[gi];
    const orExpr  = gParams.length === 1 ? gParams[0] : ('OR(' + gParams.join(',') + ')');
    return g.not ? ('NOT(' + orExpr + ')') : orExpr;
  }).join(',');

  const andExpr = parsedGroups.length === 1 ? andParts : ('AND(' + andParts + ')');

  return '=MAP(A2:A,' + ranges + ',LAMBDA(' + allParams + ',' +
         'IF(ROW(a)=2,"Filter Match",' +
         'IF(a="",' + '"",' + andExpr + '))))';
}

// Builds a MAP formula for a single term that returns TRUE/FALSE per data row,
// searching across title and abstract together (title alone stands in when
// the abstract is missing — a common gap, not a parsing bug — and cheaply
// catches title-only matches even when the abstract is present).
// Uses REGEXMATCH with \b word boundaries rather than SEARCH's plain
// substring match, so a short term like "AI" or "ML" doesn't match inside
// unrelated words ("said", "html") — mirrors jsMatchesFilter in crawl.js.
// Header row (ROW=2) returns the term string itself as the column label.
function buildTermFormula(term, titleColNum, abstractColNum) {
  const titleLetter = colToLetter(titleColNum);
  const absLetter   = colToLetter(abstractColNum);
  const safeTerm    = term.replace(/"/g, '""'); // for the header label (row 2)
  const regexTerm   = term.toLowerCase()
    .replace(/[.*+?^${}()|[\]\\]/g, '\\$&') // escape regex metacharacters
    .replace(/"/g, '""');                    // escape quotes for the formula string literal
  return '=MAP(A2:A,' + titleLetter + '2:' + titleLetter + ',' + absLetter + '2:' + absLetter + ',LAMBDA(a,t,k,' +
         'IF(ROW(a)=2,"' + safeTerm + '",' +
         'IF(a="","",' +
         'REGEXMATCH(LOWER(t&" "&k),"\\b' + regexTerm + '\\b")' +
         '))))';
}

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

function getOpenAlexEmail() {
  return PropertiesService.getScriptProperties().getProperty('OPENALEX_EMAIL') || '';
}

function getSheetName(direction) {
  return direction === "forward" ? "Snowball Forward" : "Snowball Backward";
}

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------

function getRoundCounter(sheet) {
  return sheet.getRange(1, 1).getValue() || 0;
}

function setRoundCounter(sheet, value) {
  sheet.getRange(1, 1).setValue(value);
}

function getOrCreateSnowballSheet(direction) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = getSheetName(direction);
  let sheet  = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    setupSheet(sheet);
  }
  return sheet;
}

function setupSheet(sheet) {
  sheet.clear().clearFormats();

  // Row 1: round counter
  sheet.getRange(1, 1).setValue(0).setFontWeight("bold").setFontSize(12);
  sheet.getRange(1, 2).setValue("Current Round").setFontColor("#888888").setFontStyle("italic");

  // Row 2: static headers for cols 1–13, MAP formula in col 14
  const headerStyle = { fontWeight: "bold", background: "#4285f4", fontColor: "white" };
  sheet.getRange(2, 1, 1, NUM_COLS - 1)
    .setValues([HEADERS.slice(0, -1)])
    .setFontWeight(headerStyle.fontWeight)
    .setBackground(headerStyle.background)
    .setFontColor(headerStyle.fontColor);
  sheet.getRange(2, COL.FILTER_MATCH)
    .setFormula(DEFAULT_FILTER_FORMULA)
    .setFontWeight(headerStyle.fontWeight)
    .setBackground(headerStyle.background)
    .setFontColor(headerStyle.fontColor);

  sheet.setFrozenRows(2);

  // Column widths
  sheet.setColumnWidth(COL.ROUND,     60);
  sheet.setColumnWidth(COL.INCLUDE,   70);
  sheet.setColumnWidth(COL.FETCHED,   70);
  sheet.setColumnWidth(COL.REF_COUNT, 90);
  sheet.setColumnWidth(COL.CITED_BY,  80);
  sheet.setColumnWidth(COL.YEAR,      60);
  sheet.setColumnWidth(COL.TITLE,     250);
  sheet.setColumnWidth(COL.AUTHORS,   180);
  sheet.setColumnWidth(COL.PUB_TYPES, 130);
  sheet.setColumnWidth(COL.VENUE,     130);
  sheet.setColumnWidth(COL.ABSTRACT,  400);
  sheet.setColumnWidth(COL.ID,        130);
  sheet.setColumnWidth(COL.CITES,         200);
  sheet.setColumnWidth(COL.FILTER_MATCH,  110);

  // Word wrap for all cells, then clip the abstract column
  sheet.getRange(1, 1, sheet.getMaxRows(), NUM_COLS).setWrap(true);
  sheet.getRange(1, COL.ABSTRACT, sheet.getMaxRows(), 1)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Row height and font sizes for data rows
  sheet.setRowHeights(3, sheet.getMaxRows() - 2, ROW_HEIGHT);
  const dataRows = sheet.getMaxRows() - 2;
  sheet.getRange(3, COL.TITLE,    dataRows, 1).setFontSize(10);
  sheet.getRange(3, COL.AUTHORS,  dataRows, 1).setFontSize(10);
  sheet.getRange(3, COL.ABSTRACT, dataRows, 1).setFontSize(8);

  // Filter starting at header row, covering all columns
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, NUM_COLS).createFilter();

  // Conditional formatting: highlight entire row when Filter Match column is TRUE
  const cfRange = sheet.getRange(3, 1, sheet.getMaxRows() - 2, NUM_COLS);
  const highlightRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${FILTER_MATCH_COL_LETTER}3=TRUE`)
    .setBackground("#b7e1cd")
    .setRanges([cfRange])
    .build();
  sheet.setConditionalFormatRules([highlightRule]);
}

// Returns the last sheet row that contains real data, based on column A (Round).
// sheet.getLastRow() is unreliable because the MAP formula in col N spills "" into
// every row, making Sheets think the sheet is full.
function getLastDataRow(sheet) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < 3) return 2;
  const vals = sheet.getRange(3, COL.ROUND, maxRows - 2, 1).getValues().flat();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i] !== "" && vals[i] !== null) return i + 3;
  }
  return 2; // no data rows yet
}

function clearDataRows(sheet) {
  const lastDataRow = getLastDataRow(sheet);
  if (lastDataRow >= 3) {
    sheet.getRange(3, 1, lastDataRow - 2, NUM_COLS).clearContent();
  }
}

function getExistingIds(sheet) {
  const lastDataRow = getLastDataRow(sheet);
  if (lastDataRow < 3) return new Set();
  const ids = sheet.getRange(3, COL.ID, lastDataRow - 2, 1).getValues().flat();
  return new Set(ids.filter(Boolean));
}

function applyRowFormatting(sheet, startRow, numRows) {
  sheet.setRowHeights(startRow, numRows, ROW_HEIGHT);
  sheet.getRange(startRow, 1, numRows, NUM_COLS).setWrap(true);
  sheet.getRange(startRow, COL.ABSTRACT, numRows, 1)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(startRow, COL.TITLE,    numRows, 1).setFontSize(10);
  sheet.getRange(startRow, COL.AUTHORS,  numRows, 1).setFontSize(10);
  sheet.getRange(startRow, COL.ABSTRACT, numRows, 1).setFontSize(8);
}

// ---------------------------------------------------------------------------
// Row builders
// ---------------------------------------------------------------------------

function openAlexWorkToRow(work, round, include) {
  const id       = work.id.replace("https://openalex.org/", "");
  const title    = work.title || "";
  const year     = work.publication_year || "";
  const authors  = (work.authorships || []).map(a => a.author.display_name).join("; ");
  const refCount = work.referenced_works_count || 0;
  const citedBy  = work.cited_by_count || 0;
  // Round | Include | Fetched | Ref Count | Cited By | Year | Title | Authors | Type | Venue | Abstract | ID | Cites | Filter Match
  return [round, include, false, refCount, citedBy, year, title, authors, "", "", "", id, "", ""];
}

function s2PaperToRow(paper, round) {
  const mag      = paper.externalIds && paper.externalIds.MAG;
  const id       = mag ? `W${mag}` : `S2:${paper.paperId}`;
  const title    = paper.title || "";
  const year     = paper.year || "";
  const authors  = (paper.authors || []).map(a => a.name).join("; ");
  const refCount = paper.referenceCount || 0;
  const citedBy  = paper.citationCount  || 0;
  const pubTypes = (paper.publicationTypes || []).join(", ");
  const venue    = paper.venue || "";
  const abstract = paper.abstract || "";
  // Round | Include | Fetched | Ref Count | Cited By | Year | Title | Authors | Type | Venue | Abstract | ID | Cites | Filter Match
  return [round, false, false, refCount, citedBy, year, title, authors, pubTypes, venue, abstract, id, "", ""];
}

// ---------------------------------------------------------------------------
// OpenAlex resolution (seed paper)
// ---------------------------------------------------------------------------

function resolveToOpenAlexWork(input) {
  const clean = input.trim();
  const upper = clean.toUpperCase();
  let url;

  if (upper.startsWith("HTTPS://OPENALEX.ORG/W") || upper.startsWith("W")) {
    const id = clean.replace(/https:\/\/openalex\.org\//i, "").toUpperCase();
    url = `https://api.openalex.org/works/${id}?mailto=${getOpenAlexEmail()}`;
  } else if (clean.includes("10.")) {
    const doi = clean.replace("https://doi.org/", "");
    url = `https://api.openalex.org/works/https://doi.org/${encodeURIComponent(doi)}?mailto=${getOpenAlexEmail()}`;
  } else {
    return null;
  }

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code     = response.getResponseCode();
  const text     = response.getContentText();
  if (code !== 200) throw new Error(`OpenAlex returned HTTP ${code} for ${url.split('?')[0]}`);
  const data = JSON.parse(text);
  return data.results ? data.results[0] : data;
}

// ---------------------------------------------------------------------------
// Semantic Scholar — forward citations
// ---------------------------------------------------------------------------

function s2ResolvePaperId(openAlexId, title) {
  // If already an S2 paper ID (stored from a previous forward round), use it directly
  if (openAlexId.startsWith("S2:")) return openAlexId.slice(3);

  // Strategy 1: MAG ID (works for papers indexed before MAG shutdown, ~pre-2022)
  const magId      = openAlexId.replace(/^W/i, "");
  const magUrl     = `https://api.semanticscholar.org/graph/v1/paper/MAG:${magId}?fields=paperId`;
  const magResult  = JSON.parse(s2Fetch(magUrl).getContentText());
  if (!magResult.error && !magResult.message && magResult.paperId) return `MAG:${magId}`;

  // Strategy 2: fetch DOI from OpenAlex (free direct lookup), then S2 DOI lookup
  const oaUrl    = `https://api.openalex.org/works/${openAlexId}?select=id,doi&mailto=${getOpenAlexEmail()}`;
  const oaWork   = JSON.parse(UrlFetchApp.fetch(oaUrl, { muteHttpExceptions: true }).getContentText());
  if (oaWork.doi) {
    const doi       = oaWork.doi.replace("https://doi.org/", "");
    const doiUrl    = `https://api.semanticscholar.org/graph/v1/paper/DOI:${doi}?fields=paperId`;
    const doiResult = JSON.parse(s2Fetch(doiUrl).getContentText());
    if (!doiResult.error && !doiResult.message && doiResult.paperId) return `DOI:${doi}`;
  }

  // Strategy 3: title search on Semantic Scholar
  if (title) {
    const searchUrl    = `https://api.semanticscholar.org/graph/v1/paper/search?query=${encodeURIComponent(title)}&fields=paperId&limit=1`;
    const searchResult = JSON.parse(s2Fetch(searchUrl).getContentText());
    const hit          = searchResult.data && searchResult.data[0];
    if (hit && hit.paperId) return hit.paperId;
  }

  return null;
}

function s2GetCitations(openAlexId, title) {
  const s2Id = s2ResolvePaperId(openAlexId, title);
  if (!s2Id) throw new Error(`Could not find ${openAlexId} in Semantic Scholar via MAG ID, DOI, or title search.`);

  const fields = "paperId,externalIds,title,abstract,year,authors,citationCount,referenceCount,publicationTypes,venue";
  const limit  = 500;
  let offset   = 0;
  const papers = [];

  while (true) {
    // Pace requests — stay within S2 rate limits (1.1 s between pages)
    if (offset > 0) Utilities.sleep(1100);

    const url      = `https://api.semanticscholar.org/graph/v1/paper/${s2Id}/citations?fields=${fields}&limit=${limit}&offset=${offset}`;
    const s2Resp   = s2Fetch(url);
    const s2Code   = s2Resp.getResponseCode();
    if (s2Code !== 200) throw new Error(`Semantic Scholar returned HTTP ${s2Code} — try again shortly.`);
    const result   = JSON.parse(s2Resp.getContentText());

    if (result.error || result.message) throw new Error(result.error || result.message);

    const data = result.data || [];
    data.forEach(item => papers.push(item.citingPaper));

    if (data.length < limit) break;
    offset += limit;
  }

  return papers;
}

// ---------------------------------------------------------------------------
// Main entry points
// ---------------------------------------------------------------------------

function startSnowball(input, direction) {
  direction = direction || "backward";

  var logRow = appendLogRow('Snowball', {
    name:  input,
    seeds: [input],
    runBackward:   direction === 'backward',
    expandBackward: false
  });

  try {
    const work = resolveToOpenAlexWork(input);
    if (!work || !work.id) {
      updateLogRow(logRow, 'Error: paper not found');
      return "Could not find a paper. Please enter a DOI or OpenAlex ID (e.g. W2741809807).";
    }

    const sheet = getOrCreateSnowballSheet(direction);
    clearDataRows(sheet);

    const row = openAlexWorkToRow(work, 1, false);
    sheet.getRange(3, 1, 1, NUM_COLS).setValues([row]);
    sheet.getRange(3, COL.INCLUDE).insertCheckboxes();
    sheet.getRange(3, COL.INCLUDE).setValue(true);
    sheet.getRange(3, COL.FETCHED).insertCheckboxes();
    sheet.getRange(3, COL.FETCHED).setValue(false);
    applyRowFormatting(sheet, 3, 1);

    setRoundCounter(sheet, 1);
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);

    // Update log: name with resolved title, mark complete (seed loaded)
    if (logRow) {
      const logSheet = getLogSheet();
      if (logSheet) logSheet.getRange(logRow, LOG_EXT.NAME).setValue(work.title || input);
      updateLogRow(logRow, 'Complete');
    }

    return `Seed added to "${getSheetName(direction)}" (Round 1): "${work.title}"`;
  } catch (e) {
    updateLogRow(logRow, 'Error: ' + e.message.slice(0, 50));
    return `Error: ${e.message}`;
  }
}

function expandSnowball(direction) {
  direction = direction || "backward";

  try {
    const sheet       = getOrCreateSnowballSheet(direction);
    const lastDataRow = getLastDataRow(sheet);
    if (lastDataRow < 3) return "No papers in the sheet yet. Start with a seed paper first.";

    const currentRound = getRoundCounter(sheet);
    const data = sheet.getRange(3, 1, lastDataRow - 2, NUM_COLS).getValues();

    // Papers where Include=TRUE and Fetched=FALSE
    const toExpand = data
      .map((r, i) => ({ row: r, sheetRow: i + 3 }))
      .filter(({ row }) => row[COL.INCLUDE - 1] === true && row[COL.FETCHED - 1] === false);

    if (toExpand.length === 0) {
      return "No papers to expand — either none are checked, or all checked papers have already been fetched.";
    }

    const existingIds = getExistingIds(sheet);
    const nextRound   = currentRound + 1;
    const allNewIds   = new Set();
    const s2Cache     = {};

    // ------------------------------------------------------------------
    // Collect new IDs
    // ------------------------------------------------------------------
    if (direction === "backward") {
      for (const { row, sheetRow } of toExpand) {
        const workId  = row[COL.ID - 1];
        const url     = `https://api.openalex.org/works/${workId}?mailto=${getOpenAlexEmail()}&select=id,referenced_works`;
        const resp    = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const work    = JSON.parse(resp.getContentText());
        if (work.error || work.message) {
          return `OpenAlex error fetching references for ${workId}: ${work.error || work.message}`;
        }
        const refs = work.referenced_works || [];
        if (refs.length === 0) {
          return `No references found for ${workId} — OpenAlex may not have parsed its reference list yet.`;
        }
        refs.map(r => r.replace("https://openalex.org/", ""))
            .forEach(r => { if (!existingIds.has(r)) allNewIds.add(r); });
        sheet.getRange(sheetRow, COL.FETCHED).setValue(true);
      }
    } else {
      for (const { row, sheetRow } of toExpand) {
        const workId = row[COL.ID - 1];
        const title  = row[COL.TITLE - 1];
        let papers;
        try {
          papers = s2GetCitations(workId, title);
        } catch (e) {
          return `Semantic Scholar error for ${workId}: ${e.message}`;
        }
        for (const paper of papers) {
          const mag = paper.externalIds && paper.externalIds.MAG;
          const id  = mag ? `W${mag}` : `S2:${paper.paperId}`;
          if (!existingIds.has(id)) {
            allNewIds.add(id);
            s2Cache[id] = paper;
          }
        }
        sheet.getRange(sheetRow, COL.FETCHED).setValue(true);
      }
    }

    if (allNewIds.size === 0) return "No new papers found — all are already in the sheet.";

    // ------------------------------------------------------------------
    // Build rows
    // ------------------------------------------------------------------
    const newRows = [];

    if (direction === "backward") {
      const idArray = [...allNewIds];
      for (let i = 0; i < idArray.length; i += 50) {
        const batch  = idArray.slice(i, i + 50);
        const url    = `https://api.openalex.org/works?filter=openalex_id:${batch.join("%7C")}&per_page=50&select=id,title,publication_year,authorships,referenced_works_count,cited_by_count&mailto=${getOpenAlexEmail()}`;
        const resp   = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const result = JSON.parse(resp.getContentText());
        if (result.error || result.message) {
          return `OpenAlex batch error (batch ${Math.floor(i/50)+1}): ${result.error || result.message}`;
        }
        for (const work of (result.results || [])) {
          newRows.push(openAlexWorkToRow(work, nextRound, false));
          existingIds.add(work.id.replace("https://openalex.org/", ""));
        }
      }
    } else {
      for (const id of allNewIds) {
        const paper = s2Cache[id];
        if (paper) {
          newRows.push(s2PaperToRow(paper, nextRound));
          existingIds.add(id);
        }
      }
    }

    if (newRows.length === 0) return "No metadata could be retrieved for the new papers.";

    // ------------------------------------------------------------------
    // Write rows to sheet
    // ------------------------------------------------------------------
    const startRow = getLastDataRow(sheet) + 1;
    sheet.getRange(startRow, 1, newRows.length, NUM_COLS).setValues(newRows);
    sheet.getRange(startRow, COL.INCLUDE, newRows.length, 1).insertCheckboxes();
    sheet.getRange(startRow, COL.FETCHED,  newRows.length, 1).insertCheckboxes();
    applyRowFormatting(sheet, startRow, newRows.length);

    setRoundCounter(sheet, nextRound);

    if (direction === "backward") updateCitesColumn(sheet);

    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
    return `Round ${nextRound} complete: ${newRows.length} new papers added.`;
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

// ---------------------------------------------------------------------------
// Highlight filter — parses boolean query → conditional formatting rule
// ---------------------------------------------------------------------------

// groups = [{ terms: "term A, term B, term C", not: false }, ...]
// Each group is AND'd together; terms within a group are OR'd.
function getLastSnowballFilter() {
  const stored = PropertiesService.getScriptProperties().getProperty('SNOWBALL_FILTER_GROUPS');
  if (!stored) return null;
  try { return JSON.parse(stored); } catch(e) { return null; }
}

function applySnowballFilter(groups, direction) {
  try {
    PropertiesService.getScriptProperties()
      .setProperty('SNOWBALL_FILTER_GROUPS', JSON.stringify(groups));

    const sheet = getOrCreateSnowballSheet(direction);

    // Parse terms from each group, drop empty groups
    const parsed = groups.map(function(g) {
      return {
        not:   g.not,
        terms: (g.terms || '').split(',')
          .map(function(t) { return t.trim().replace(/^["'"']|["'"']$/g, '').trim(); })
          .filter(function(t) { return t.length > 0; })
      };
    }).filter(function(g) { return g.terms.length > 0; });

    if (parsed.length === 0) return 'No valid terms found — check your groups.';

    // --- Per-term helper columns ---
    // Clear previous helper area first.
    const maxHelper = 60;
    sheet.getRange(1, FIRST_DETAIL_COL, 1, maxHelper).breakApart().clearContent().clearFormat();
    sheet.getRange(2, FIRST_DETAIL_COL, 1, maxHelper).clearContent().clearFormat();

    let colNum = FIRST_DETAIL_COL;
    parsed.forEach(function(g, groupIdx) {
      const groupStartCol = colNum;
      const isNot  = g.not;
      const termBg = isNot ? '#fce8e6' : '#e8f0fe';

      g.terms.forEach(function(term) {
        sheet.getRange(2, colNum)
          .setFormula(buildTermFormula(term, COL.TITLE, COL.ABSTRACT))
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
        const label = (isNot ? 'NOT Group ' : 'Group ') + (groupIdx + 1);
        const hdrBg = isNot ? '#e53935' : '#1a73e8';
        sheet.getRange(1, groupStartCol, 1, g.terms.length)
          .merge()
          .setValue(label)
          .setBackground(hdrBg)
          .setFontColor('white')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');

        // Thick left border separating groups
        sheet.getRange(1, groupStartCol, sheet.getMaxRows(), 1)
          .setBorder(null, true, null, null, null, null,
                     '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    });

    // Tall row 2 for rotated labels + CF rules for helper columns
    if (colNum > FIRST_DETAIL_COL) {
      sheet.setRowHeight(2, 130);

      // CF: TRUE cells → yellow block (both background and font #FFF176, so value is hidden)
      //     FALSE cells → white on white (invisible)
      // This replaces the old checkbox-validation approach which blocked MAP spills.
      const termDataRange = sheet.getRange(3, FIRST_DETAIL_COL, sheet.getMaxRows() - 2, colNum - FIRST_DETAIL_COL);
      const startLetter = colToLetter(FIRST_DETAIL_COL);
      const trueRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + startLetter + '3=TRUE')
        .setBackground('#FFF176').setFontColor('#FFF176')
        .setRanges([termDataRange]).build();
      const falseRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=' + startLetter + '3=FALSE')
        .setBackground('#FFFFFF').setFontColor('#FFFFFF')
        .setRanges([termDataRange]).build();
      // Preserve existing rules (e.g. the green row-highlight rule) but drop any
      // that targeted the helper-column area so stale rules don't accumulate.
      const existingRules = sheet.getConditionalFormatRules().filter(function(rule) {
        return !rule.getRanges().some(function(r) { return r.getColumn() >= FIRST_DETAIL_COL; });
      });
      sheet.setConditionalFormatRules(existingRules.concat([trueRule, falseRule]));
    }

    // --- Filter Match: references per-term columns so logic isn't duplicated ---
    const filterFormula = buildFilterMatchFromTermCols(parsed, FIRST_DETAIL_COL);
    if (!filterFormula) return 'No valid terms found — check your groups.';
    sheet.getRange(2, COL.FILTER_MATCH)
      .setFormula(filterFormula)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white');

    return 'Highlight rule applied to "' + getSheetName(direction) + '".';
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

// Builds the inner logical expression using lambda variable `k` for the abstract cell.
// Groups are AND'd together; terms within each group are OR'd.
function buildInnerExpression(groups) {
  const andParts = [];

  for (const group of groups) {
    const terms = group.terms
      .split(",")
      .map(t => t.trim().replace(/^["'""]|["'""]$/g, "").trim())
      .filter(t => t.length > 0);

    if (terms.length === 0) continue;

    const searches = terms.map(t => `ISNUMBER(SEARCH("${t}",k))`);
    const orPart   = searches.length === 1 ? searches[0] : `OR(${searches.join(",")})`;
    andParts.push(group.not ? `NOT(${orPart})` : orPart);
  }

  if (andParts.length === 0) return null;
  return andParts.length === 1 ? andParts[0] : `AND(${andParts.join(",")})`;
}

// ---------------------------------------------------------------------------
// Cites column (backward only — uses OpenAlex direct lookups)
// ---------------------------------------------------------------------------

function updateCitesColumn(sheet) {
  const lastRow = getLastDataRow(sheet);
  if (lastRow < 3) return;

  const ids   = sheet.getRange(3, COL.ID, lastRow - 2, 1).getValues().flat();
  const idSet = new Set(ids.filter(Boolean));

  for (let i = 0; i < ids.length; i++) {
    const workId = ids[i];
    if (!workId || workId.startsWith("S2:")) continue;

    const url    = `https://api.openalex.org/works/${workId}?mailto=${getOpenAlexEmail()}&select=id,referenced_works`;
    const work   = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
    const refs   = (work.referenced_works || []).map(r => r.replace("https://openalex.org/", ""));
    const inList = refs.filter(r => idSet.has(r));
    sheet.getRange(i + 3, COL.CITES).setValue(inList.join(", "));
  }
}
